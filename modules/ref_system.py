"""Reference number system: auto-incremental EDU-PRO-YY-SEQ format."""

import json
from datetime import datetime

from config import DATA_DIR, REF_PREFIX

PROPOSALS_FILE = DATA_DIR / "proposals.json"


def _load_proposals() -> dict:
    """Load the proposals registry."""
    if PROPOSALS_FILE.exists():
        return json.loads(PROPOSALS_FILE.read_text(encoding="utf-8"))
    return {"proposals": [], "last_seq": {}}


def _save_proposals(data: dict):
    """Save the proposals registry."""
    DATA_DIR.mkdir(exist_ok=True)
    PROPOSALS_FILE.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")


def generate_ref_number() -> str:
    """Generate the next reference number: EDU-PRO-YY-SEQ."""
    data = _load_proposals()
    year = datetime.now().strftime("%y")  # e.g., "26"

    last_seq = data.get("last_seq", {})
    seq = last_seq.get(year, 0) + 1
    last_seq[year] = seq
    data["last_seq"] = last_seq

    ref = f"{REF_PREFIX}-{year}-{seq:03d}"
    _save_proposals(data)
    return ref


def register_proposal(ref_number: str, client: str, project: str, mode: str):
    """Register a generated proposal in the registry."""
    data = _load_proposals()
    data["proposals"].append({
        "ref": ref_number,
        "client": client,
        "project": project,
        "mode": mode,
        "date": datetime.now().isoformat(),
    })
    _save_proposals(data)
