"""Knowledge base loader: reads markdown files from the knowledge/ directory."""

from config import KNOWLEDGE_DIR

_cache: dict[str, str] = {}


def load_knowledge() -> dict[str, str]:
    """Load all .md files from the knowledge directory. Returns dict of filename -> content."""
    if _cache:
        return _cache

    for md_file in sorted(KNOWLEDGE_DIR.glob("*.md")):
        _cache[md_file.stem] = md_file.read_text(encoding="utf-8")

    return _cache


def get_knowledge_context() -> str:
    """Get all knowledge concatenated as a single context string for the AI prompt."""
    knowledge = load_knowledge()
    sections = []
    for name, content in knowledge.items():
        sections.append(f"--- {name} ---\n{content}")
    return "\n\n".join(sections)
