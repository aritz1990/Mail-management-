#!/usr/bin/env python3
"""
Attio CRM API client for the pitch deck processor.
IMPORTANT: This module never deletes any Attio records or data.
"""

import os
import requests

BASE_URL = "https://api.attio.com/v2"
ANNA_RITZ_EMAIL = "ar@angelinvest.ventures"

# Discovered at runtime
_pitch_deck_slug = None
_anna_ritz_member_id = None
_workspace_slug = None


def _headers():
    key = os.environ["ATTIO_API_KEY"]
    return {"Authorization": f"Bearer {key}", "Content-Type": "application/json"}


def get_record_url(record_id: str) -> str:
    """Return a direct link to the Attio company record."""
    slug = _workspace_slug or "angelinvest"
    return f"https://app.attio.com/{slug}/company/{record_id}/overview"


def initialise():
    """Discover field slugs, workspace slug, and Anna Ritz's member ID. Call once at startup."""
    global _pitch_deck_slug, _anna_ritz_member_id, _workspace_slug

    # Discover workspace slug
    r = requests.get(f"{BASE_URL}/workspace", headers=_headers())
    if r.status_code == 200:
        _workspace_slug = r.json().get("data", {}).get("slug")
    print(f"  Attio workspace slug: {_workspace_slug}")

    # Discover pitch deck URL slug
    r = requests.get(f"{BASE_URL}/objects/companies/attributes", headers=_headers())
    r.raise_for_status()
    for attr in r.json().get("data", []):
        title = attr.get("title", "").lower()
        if "pitch" in title and ("deck" in title or "url" in title):
            _pitch_deck_slug = attr["api_slug"]
            break
    if not _pitch_deck_slug:
        _pitch_deck_slug = "pitch_deck_url"
    print(f"  Attio pitch deck slug: {_pitch_deck_slug}")

    # Discover Anna Ritz's member ID
    r = requests.get(f"{BASE_URL}/workspace_members", headers=_headers())
    r.raise_for_status()
    members_data = r.json().get("data", [])
    print(f"  Attio workspace members found: {len(members_data)}")
    if members_data:
        print(f"  First member sample: {members_data[0]}")
    for member in members_data:
        addr = member.get("email_address", "")
        if addr.lower() == ANNA_RITZ_EMAIL.lower():
            _anna_ritz_member_id = member["id"]["workspace_member_id"]
            break
    print(f"  Anna Ritz member ID: {_anna_ritz_member_id}")


def _search(filters: dict) -> list:
    r = requests.post(
        f"{BASE_URL}/objects/companies/records/query",
        headers=_headers(),
        json={"filter": filters}
    )
    r.raise_for_status()
    return r.json().get("data", [])


def match_company(company_name: str, domain: str | None, alt_name: str | None) -> tuple[str, list]:
    """
    Match a company in Attio using domain, then exact name, then partial name.

    Returns (status, candidates):
      - 'single_match': one confident match found
      - 'no_match':     no candidates found
      - 'ambiguous':    multiple or fuzzy candidates, needs human review
    """
    # 1. Domain match — most reliable
    if domain:
        results = _search({"domains": {"$contains": domain}})
        if len(results) == 1:
            return "single_match", results
        if len(results) > 1:
            return "ambiguous", results

    # 2. Exact name match on company name and alt name
    for name in filter(None, [company_name, alt_name]):
        results = _search({"name": {"$eq": name}})
        if len(results) == 1:
            return "single_match", results
        if len(results) > 1:
            return "ambiguous", results

    # 3. Partial name match — catches "Sprive" vs "Sprive Ltd"
    for name in filter(None, [company_name, alt_name]):
        results = _search({"name": {"$contains": name}})
        if results:
            return "ambiguous", results

    return "no_match", []


def get_record_id(record: dict) -> str:
    return record["id"]["record_id"]


def get_company_name(record: dict) -> str:
    values = record.get("values", {}).get("name", [])
    return values[0].get("value", "") if values else ""


def is_owned_by_anna_ritz(record: dict) -> bool:
    owner_values = record.get("values", {}).get("owner", [])
    if not owner_values:
        return False
    owner_id = owner_values[0].get("referenced_actor_id")
    return owner_id == _anna_ritz_member_id


def update_pitch_deck_url(record_id: str, url: str):
    """Update the Pitch deck URL field on an existing company. Never deletes anything."""
    r = requests.patch(
        f"{BASE_URL}/objects/companies/records/{record_id}",
        headers=_headers(),
        json={"data": {"values": {_pitch_deck_slug: [{"value": url}]}}}
    )
    r.raise_for_status()
    return r.json()


def create_company(name: str, domain: str | None, pitch_deck_url: str) -> dict:
    """Create a new company record with pitch deck URL pre-filled."""
    values = {
        "name": [{"value": name}],
        _pitch_deck_slug: [{"value": pitch_deck_url}],
    }
    if domain:
        values["domains"] = [{"domain": domain}]
    r = requests.post(
        f"{BASE_URL}/objects/companies/records",
        headers=_headers(),
        json={"data": {"values": values}}
    )
    r.raise_for_status()
    return r.json()
