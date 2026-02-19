"""Data types for El Licitador."""

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class PricingItem:
    category: str
    product: str
    quantity: int
    unit_price: float

    @property
    def subtotal(self) -> float:
        return self.quantity * self.unit_price


@dataclass
class ProposalContent:
    """AI-generated content for the proposal."""
    introduction: str = ""
    company_presentation: str = ""
    technical_proposal: Optional[str] = None
    deliverables: Optional[str] = None
    conclusion: str = ""


@dataclass
class ProposalData:
    """All data needed to generate a proposal document."""
    # Metadata
    date: str
    ref_number: str
    project_title: str
    client_name: str

    # Mode
    is_formal: bool = True

    # Content (AI-generated or manual)
    content: ProposalContent = field(default_factory=ProposalContent)

    # Pricing
    pricing_items: list[PricingItem] = field(default_factory=list)

    # Commercial terms
    commercial_terms: list[str] = field(default_factory=list)

    @property
    def total(self) -> float:
        return sum(item.subtotal for item in self.pricing_items)
