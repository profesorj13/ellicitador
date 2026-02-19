"""AI engine: generates proposal content using the Anthropic Claude API."""

import json
import anthropic

from modules.knowledge_base import get_knowledge_context
from modules.types import ProposalContent


SYSTEM_PROMPT = """Sos un escritor experto de propuestas comerciales para Educabot, una empresa argentina de tecnología educativa. Escribís en español argentino, con tono profesional pero cálido.

Tu tarea es generar el contenido textual de una propuesta comercial. NO generás precios ni condiciones comerciales (eso viene por separado). Solo generás el texto narrativo.

Reglas de escritura:
- Usá español argentino formal (no voseo excesivo, pero natural)
- Tono profesional y cálido, como si hablaras con un cliente importante
- Mencioná siempre que Educabot tiene "más de diez años de experiencia"
- Sé concreto y específico, no genérico
- Los párrafos deben ser de 2-4 oraciones
- Separá párrafos con doble salto de línea

Base de conocimiento de Educabot:
{knowledge_context}"""


FORMAL_PROMPT = """Generá el contenido para una propuesta comercial FORMAL (aún no vendido, necesita convencer).

Datos:
- Cliente: {client_name}
- Proyecto: {project_title}
- Productos/servicios a incluir: {products}
- Contexto adicional: {context}

Generá un JSON con esta estructura exacta:
{{
  "introduction": "2-3 párrafos presentando a Educabot y el propósito de la propuesta. Separá párrafos con \\n\\n",
  "company_presentation": "2-3 párrafos sobre Educabot, su experiencia, trayectoria y certificaciones. Separá párrafos con \\n\\n",
  "technical_proposal": "3-5 párrafos describiendo la solución técnica propuesta, incluyendo propósito, audiencia y detalle. Separá párrafos con \\n\\n",
  "deliverables": "Lista de entregables agrupados: Digitales, Físicos y Servicios. Separá grupos con \\n\\n",
  "conclusion": "4 párrafos: agradecimiento, próximos pasos, compromiso de acompañamiento, despedida cordial. Separá párrafos con \\n\\n"
}}

Respondé SOLO con el JSON, sin texto adicional."""


INFORMAL_PROMPT = """Generá el contenido para una propuesta comercial INFORMAL (ya vendido, solo formalizar).

Datos:
- Cliente: {client_name}
- Proyecto: {project_title}
- Productos/servicios a incluir: {products}
- Contexto adicional: {context}

Generá un JSON con esta estructura exacta:
{{
  "introduction": "1 párrafo breve presentando la propuesta y su contexto",
  "company_presentation": "",
  "technical_proposal": null,
  "deliverables": null,
  "conclusion": "1-2 párrafos breves de cierre cordial con datos de contacto"
}}

Respondé SOLO con el JSON, sin texto adicional."""


def generate_proposal_content(
    client_name: str,
    project_title: str,
    products: list[str],
    context: str = "",
    is_formal: bool = True,
    api_key: str | None = None,
) -> ProposalContent:
    """Generate proposal content using Claude API.

    Args:
        api_key: Anthropic API key. Falls back to ANTHROPIC_API_KEY env var if not provided.

    Returns a ProposalContent dataclass with the generated text sections.
    """
    client = anthropic.Anthropic(api_key=api_key) if api_key else anthropic.Anthropic()
    knowledge_context = get_knowledge_context()

    system = SYSTEM_PROMPT.format(knowledge_context=knowledge_context)
    template = FORMAL_PROMPT if is_formal else INFORMAL_PROMPT
    user_prompt = template.format(
        client_name=client_name,
        project_title=project_title,
        products=", ".join(products) if products else "A definir según contexto",
        context=context or "Sin contexto adicional",
    )

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=4096,
        system=system,
        messages=[{"role": "user", "content": user_prompt}],
    )

    # Parse the JSON response
    response_text = response.content[0].text.strip()

    # Handle potential markdown code blocks wrapping the JSON
    if response_text.startswith("```"):
        lines = response_text.split("\n")
        # Remove first and last lines (```json and ```)
        response_text = "\n".join(lines[1:-1])

    data = json.loads(response_text)

    return ProposalContent(
        introduction=data.get("introduction", ""),
        company_presentation=data.get("company_presentation", ""),
        technical_proposal=data.get("technical_proposal"),
        deliverables=data.get("deliverables"),
        conclusion=data.get("conclusion", ""),
    )
