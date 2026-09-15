#!/usr/bin/env python3
"""Copy a folder between OneDrive accounts in different Microsoft 365 tenants."""

from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import time
from pathlib import Path
from typing import Any, Iterator

import msal
import requests

GRAPH_ROOT = "https://graph.microsoft.com/v1.0"
CHUNK_SIZE = 10 * 1024 * 1024
RETRYABLE_STATUS_CODES = {429, 500, 502, 503, 504}


class GraphClient:
    def __init__(self, tenant_id: str, client_id: str, client_secret: str) -> None:
        authority = f"https://login.microsoftonline.com/{tenant_id}"
        self.app = msal.ConfidentialClientApplication(
            client_id, authority=authority, client_credential=client_secret
        )
        self.session = requests.Session()
        self.token: str | None = None

    def _access_token(self) -> str:
        if self.token is None:
            result = self.app.acquire_token_for_client(
                scopes=["https://graph.microsoft.com/.default"]
            )
            if "access_token" not in result:
                raise RuntimeError(f"Could not acquire Graph token: {result.get('error_description', result)}")
            self.token = result["access_token"]
        return self.token

    def request(self, method: str, url: str, **kwargs: Any) -> requests.Response:
        if not url.startswith("http"):
            url = GRAPH_ROOT + url
        headers = kwargs.pop("headers", {})
        headers["Authorization"] = f"Bearer {self._access_token()}"
        for attempt in range(6):
            response = self.session.request(method, url, headers=headers, timeout=120, **kwargs)
            if response.status_code not in RETRYABLE_STATUS_CODES:
                break
            delay = int(response.headers.get("Retry-After", min(2**attempt, 60)))
            logging.warning("Graph returned %s; retrying in %ss", response.status_code, delay)
            time.sleep(delay)
        if response.status_code >= 400:
            raise RuntimeError(f"Graph {method} {url} failed ({response.status_code}): {response.text[:500]}")
        return response

    def get_item_by_path(self, user: str, path: str) -> dict[str, Any]:
        encoded_path = "/".join(requests.utils.quote(part, safe="") for part in path.strip("/").split("/"))
        suffix = f"/users/{requests.utils.quote(user, safe='')}/drive/root"
        if encoded_path:
            suffix += f":/{encoded_path}:"
        return self.request("GET", suffix).json()

    def ensure_folder_path(self, user: str, path: str) -> dict[str, Any]:
        root = self.get_item_by_path(user, "")
        current = root
        for name in filter(None, path.strip("/").split("/")):
            existing = self.find_child(user, current["id"], name)
            if existing and "folder" not in existing:
                raise RuntimeError(f"Destination item named {name!r} is not a folder")
            current = existing or self.create_folder(user, current["id"], name)
        return current

    def children(self, user: str, item_id: str) -> Iterator[dict[str, Any]]:
        url = f"/users/{requests.utils.quote(user, safe='')}/drive/items/{item_id}/children"
        while url:
            page = self.request("GET", url, params={"$top": 200}).json()
            yield from page.get("value", [])
            url = page.get("@odata.nextLink")

    def create_folder(self, user: str, parent_id: str, name: str) -> dict[str, Any]:
        response = self.request(
            "POST",
            f"/users/{requests.utils.quote(user, safe='')}/drive/items/{parent_id}/children",
            json={"name": name, "folder": {}, "@microsoft.graph.conflictBehavior": "fail"},
        )
        return response.json()

    def find_child(self, user: str, parent_id: str, name: str) -> dict[str, Any] | None:
        for item in self.children(user, parent_id):
            if item.get("name") == name:
                return item
        return None

    def download(self, user: str, item_id: str) -> requests.Response:
        return self.request(
            "GET",
            f"/users/{requests.utils.quote(user, safe='')}/drive/items/{item_id}/content",
            stream=True,
        )

    def upload(self, user: str, parent_id: str, name: str, source: requests.Response, size: int | None) -> dict[str, Any]:
        session = self.request(
            "POST",
            f"/users/{requests.utils.quote(user, safe='')}/drive/items/{parent_id}:/{requests.utils.quote(name, safe='')}: /createUploadSession".replace(": /", ":/"),
            json={"item": {"@microsoft.graph.conflictBehavior": "replace", "name": name}},
        ).json()
        upload_url = session["uploadUrl"]
        offset = 0
        for chunk in source.iter_content(chunk_size=CHUNK_SIZE):
            if not chunk:
                continue
            end = offset + len(chunk) - 1
            headers = {
                "Content-Length": str(len(chunk)),
                "Content-Range": f"bytes {offset}-{end}/{size if size is not None else '*'}",
            }
            response = self.request("PUT", upload_url, headers=headers, data=chunk)
            offset = end + 1
            if response.status_code == 201 or response.status_code == 200:
                return response.json()
        raise RuntimeError(f"Upload session for {name} did not return a completed item")


def load_state(path: Path) -> dict[str, str]:
    if not path.exists():
        return {}
    return json.loads(path.read_text())


def save_state(path: Path, state: dict[str, str]) -> None:
    temporary = path.with_suffix(".tmp")
    temporary.write_text(json.dumps(state, indent=2, sort_keys=True) + "\n")
    temporary.replace(path)


def migrate_item(
    source: GraphClient,
    destination: GraphClient,
    source_user: str,
    destination_user: str,
    source_item: dict[str, Any],
    destination_parent_id: str,
    state: dict[str, str],
    state_path: Path,
    dry_run: bool,
) -> None:
    source_id = source_item["id"]
    name = source_item["name"]
    if source_id in state:
        logging.info("Skipping completed item: %s", name)
        return

    if "folder" in source_item:
        if dry_run:
            destination_folder = {"id": "dry-run"}
        else:
            existing = destination.find_child(destination_user, destination_parent_id, name)
            if existing and "folder" not in existing:
                raise RuntimeError(f"Destination item named {name!r} is not a folder")
            destination_folder = existing or destination.create_folder(destination_user, destination_parent_id, name)
        logging.info("Folder: %s", name)
        for child in source.children(source_user, source_id):
            migrate_item(source, destination, source_user, destination_user, child, destination_folder["id"], state, state_path, dry_run)
    else:
        size = source_item.get("size")
        logging.info("File: %s (%s bytes)", name, size if size is not None else "unknown")
        if not dry_run:
            with source.download(source_user, source_id) as response:
                destination.upload(destination_user, destination_parent_id, name, response, size)
    state[source_id] = name
    if not dry_run:
        save_state(state_path, state)


def required_env(name: str) -> str:
    value = os.getenv(name)
    if not value:
        raise SystemExit(f"Missing required environment variable: {name}")
    return value


def load_env_file(path: Path = Path(".env")) -> None:
    if not path.exists():
        return
    for line in path.read_text().splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        name, value = line.split("=", 1)
        os.environ.setdefault(name.strip(), value.strip().strip('"').strip("'"))


def main() -> int:
    load_env_file()
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source-user", required=True, help="Source user's UPN or object ID")
    parser.add_argument("--destination-user", required=True, help="Destination user's UPN or object ID")
    parser.add_argument("--source-folder", help="Override SOURCE_FOLDER_PATH from .env")
    parser.add_argument("--destination-folder", help="Override DESTINATION_FOLDER_PATH from .env")
    parser.add_argument("--state", type=Path, default=Path("migration-state.json"))
    parser.add_argument("--dry-run", action="store_true", help="Walk and report items without creating or uploading anything")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()
    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO, format="%(levelname)s %(message)s")

    source_folder = args.source_folder if args.source_folder is not None else required_env("SOURCE_FOLDER_PATH")
    destination_folder = args.destination_folder if args.destination_folder is not None else required_env("DESTINATION_FOLDER_PATH")

    source = GraphClient(required_env("SOURCE_TENANT_ID"), required_env("SOURCE_CLIENT_ID"), required_env("SOURCE_CLIENT_SECRET"))
    destination = GraphClient(required_env("DEST_TENANT_ID"), required_env("DEST_CLIENT_ID"), required_env("DEST_CLIENT_SECRET"))
    source_root = source.get_item_by_path(args.source_user, source_folder)
    destination_root = (
        {"id": "dry-run", "folder": {}}
        if args.dry_run
        else destination.ensure_folder_path(args.destination_user, destination_folder)
    )
    if "folder" not in source_root or "folder" not in destination_root:
        raise SystemExit("Both source and destination paths must identify folders")
    state = load_state(args.state)
    for item in source.children(args.source_user, source_root["id"]):
        migrate_item(source, destination, args.source_user, args.destination_user, item, destination_root["id"], state, args.state, args.dry_run)
    logging.info("Migration complete")
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        logging.error("Migration interrupted; rerun to resume from saved state")
        sys.exit(130)
