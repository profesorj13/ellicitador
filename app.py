"""El Licitador - Generador de Propuestas Comerciales para Educabot."""

import anthropic
import streamlit as st
import pandas as pd
from datetime import datetime

from modules.types import ProposalData, PricingItem, ProposalContent
from modules.docx_engine import generate_proposal
from modules.ref_system import generate_ref_number, register_proposal
from modules.ai_engine import generate_proposal_content

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------

st.set_page_config(
    page_title="El Licitador - Educabot",
    page_icon="📄",
    layout="wide",
)

# ---------------------------------------------------------------------------
# Auth: API key login
# ---------------------------------------------------------------------------

if "api_key" not in st.session_state:
    st.session_state.api_key = None

if not st.session_state.api_key:
    st.markdown(
        "<div style='display:flex;justify-content:center;padding-top:40px'>"
        "<div style='max-width:420px;width:100%'>",
        unsafe_allow_html=True,
    )
    st.image("assets/logo.png", use_container_width=True)
    st.markdown("### El Licitador")
    st.caption("Generador de Propuestas Comerciales")
    st.markdown("---")

    key_input = st.text_input(
        "API Key de Anthropic",
        type="password",
        placeholder="sk-ant-...",
        help="Necesitamos tu API key para generar contenido con Claude. Se guarda solo en tu sesion.",
    )

    login_clicked = st.button("Ingresar", type="primary", use_container_width=True)

    if login_clicked:
        if not key_input or not key_input.startswith("sk-ant-"):
            st.error("La API key debe comenzar con `sk-ant-`")
        else:
            # Validate key with a minimal API call
            with st.spinner("Validando API key..."):
                try:
                    test_client = anthropic.Anthropic(api_key=key_input)
                    test_client.messages.create(
                        model="claude-haiku-4-5",
                        max_tokens=10,
                        messages=[{"role": "user", "content": "hi"}],
                    )
                    st.session_state.api_key = key_input
                    st.rerun()
                except anthropic.AuthenticationError:
                    st.error("API key invalida. Verificala e intenta de nuevo.")
                except Exception as e:
                    st.error(f"Error al validar: {e}")

    st.markdown("</div></div>", unsafe_allow_html=True)
    st.stop()

# ---------------------------------------------------------------------------
# Session state initialization
# ---------------------------------------------------------------------------

if "pricing_rows" not in st.session_state:
    st.session_state.pricing_rows = pd.DataFrame({
        "Categoría": ["Hardware"],
        "Producto": [""],
        "Cantidad": [1],
        "Valor Unitario (USD)": [0.0],
    })

if "generated_docx" not in st.session_state:
    st.session_state.generated_docx = None

if "proposal_content" not in st.session_state:
    st.session_state.proposal_content = None

if "generation_status" not in st.session_state:
    st.session_state.generation_status = None

if "ref_number" not in st.session_state:
    st.session_state.ref_number = None

# ---------------------------------------------------------------------------
# Sidebar: inputs
# ---------------------------------------------------------------------------

with st.sidebar:
    st.image("assets/logo.png", use_container_width=True)
    col_title, col_logout = st.columns([3, 1])
    with col_logout:
        if st.button("Salir", type="secondary"):
            st.session_state.api_key = None
            st.rerun()
    st.markdown("---")
    st.header("Nueva Propuesta")

    client_name = st.text_input("Cliente", placeholder="Nombre del cliente")
    project_title = st.text_input("Proyecto", placeholder="Nombre del proyecto")

    mode = st.radio(
        "Tipo de propuesta",
        ["Formal (no vendido)", "Informal (ya vendido)"],
        help="Formal: documento completo. Informal: solo tabla de precios + condiciones.",
    )
    is_formal = mode.startswith("Formal")

    st.markdown("#### Productos")
    product_options = [
        "Kit Ingenia",
        "Kit IAco",
        "Creabots Iniciación",
        "Creabots Automatización",
        "Creabots Conducción",
        "Creabots IA",
        "Impresora 3D Educativa",
        "Plataforma TRAIN",
        "Plataforma Robots",
        "Plataforma Educativa (White-label)",
        "TUNI AI",
        "AlizIA",
        "Avatares Digitales IA",
        "Capacitación Docente",
        "Soporte Técnico",
        "Desarrollo a Medida",
    ]
    selected_products = st.multiselect(
        "Seleccionar productos",
        product_options,
        label_visibility="collapsed",
    )

    additional_context = st.text_area(
        "Contexto adicional",
        placeholder="Información extra para que la IA genere mejor contenido...",
        height=100,
    )

    st.markdown("---")
    st.markdown("#### Condiciones comerciales")

    default_terms = [
        "Plazo de entrega: 30 días hábiles a partir de la confirmación de la orden de compra.",
        "Capacitación: 10 días hábiles posteriores a la entrega del equipamiento.",
        "Garantía: 12 meses sobre el equipamiento entregado.",
        "Moneda: Todos los valores están expresados en dólares estadounidenses (USD).",
        "Los precios incluyen IVA.",
        "Validez de la oferta: 30 días corridos.",
    ]
    terms_text = st.text_area(
        "Condiciones (una por línea)",
        value="\n".join(default_terms),
        height=180,
    )
    commercial_terms = [t.strip() for t in terms_text.split("\n") if t.strip()]

# ---------------------------------------------------------------------------
# Main area
# ---------------------------------------------------------------------------

st.title("Propuesta Comercial")

# Pricing table editor
st.subheader("Tabla de Precios")

category_options = ["Hardware", "Software", "Servicios", "Capacitación", "Contenidos", "Otro"]

edited_df = st.data_editor(
    st.session_state.pricing_rows,
    num_rows="dynamic",
    column_config={
        "Categoría": st.column_config.SelectboxColumn(
            "Categoría",
            options=category_options,
            required=True,
        ),
        "Producto": st.column_config.TextColumn("Producto", required=True),
        "Cantidad": st.column_config.NumberColumn("Cantidad", min_value=1, step=1, required=True),
        "Valor Unitario (USD)": st.column_config.NumberColumn(
            "Valor Unitario (USD)",
            min_value=0.0,
            step=0.01,
            format="$%.2f",
            required=True,
        ),
    },
    use_container_width=True,
    key="pricing_editor",
)

# Update session state
st.session_state.pricing_rows = edited_df

# Show calculated totals
if not edited_df.empty and edited_df["Producto"].any():
    edited_df["Sub-Total"] = edited_df["Cantidad"] * edited_df["Valor Unitario (USD)"]
    total = edited_df["Sub-Total"].sum()
    st.metric("Total", f"${total:,.2f} USD")

st.markdown("---")

# ---------------------------------------------------------------------------
# Generate button
# ---------------------------------------------------------------------------

col1, col2 = st.columns([1, 3])

with col1:
    generate_button = st.button(
        "Generar Propuesta",
        type="primary",
        use_container_width=True,
        disabled=not (client_name and project_title),
    )

if not (client_name and project_title):
    st.info("Completá el nombre del cliente y del proyecto para generar la propuesta.")

if generate_button:
    # Validate inputs
    if not client_name or not project_title:
        st.error("Completá el nombre del cliente y del proyecto.")
    else:
        # Generate reference number
        ref_number = generate_ref_number()
        st.session_state.ref_number = ref_number

        # Build pricing items from dataframe
        pricing_items = []
        for _, row in edited_df.iterrows():
            if row["Producto"] and row["Producto"].strip():
                pricing_items.append(PricingItem(
                    category=row["Categoría"],
                    product=row["Producto"],
                    quantity=int(row["Cantidad"]),
                    unit_price=float(row["Valor Unitario (USD)"]),
                ))

        # Generate AI content (if formal, or for intro/conclusion)
        with st.spinner("Generando contenido con IA..."):
            try:
                content = generate_proposal_content(
                    client_name=client_name,
                    project_title=project_title,
                    products=selected_products,
                    context=additional_context,
                    is_formal=is_formal,
                    api_key=st.session_state.api_key,
                )
                st.session_state.proposal_content = content
            except Exception as e:
                st.error(f"Error al generar contenido: {e}")
                # Fallback content
                content = ProposalContent(
                    introduction=f"Nos es grato presentar ante {client_name} la siguiente propuesta comercial para {project_title}.",
                    conclusion="Quedamos a su disposición para ampliar cualquier aspecto de la presente propuesta.",
                )
                st.session_state.proposal_content = content

        # Format date in Spanish
        months = {
            1: "enero", 2: "febrero", 3: "marzo", 4: "abril",
            5: "mayo", 6: "junio", 7: "julio", 8: "agosto",
            9: "septiembre", 10: "octubre", 11: "noviembre", 12: "diciembre",
        }
        now = datetime.now()
        date_str = f"{now.day} de {months[now.month]} de {now.year}"

        # Build ProposalData
        proposal_data = ProposalData(
            date=date_str,
            ref_number=ref_number,
            project_title=project_title,
            client_name=client_name,
            is_formal=is_formal,
            content=st.session_state.proposal_content,
            pricing_items=pricing_items,
            commercial_terms=commercial_terms,
        )

        # Generate DOCX
        with st.spinner("Generando documento .docx..."):
            docx_bytes = generate_proposal(proposal_data)
            st.session_state.generated_docx = docx_bytes

        # Register proposal
        register_proposal(ref_number, client_name, project_title,
                         "formal" if is_formal else "informal")

        st.session_state.generation_status = "success"
        st.rerun()

# ---------------------------------------------------------------------------
# Show results
# ---------------------------------------------------------------------------

if st.session_state.generation_status == "success":
    st.success(f"Propuesta generada exitosamente. Ref: **{st.session_state.ref_number}**")

    # Preview content
    if st.session_state.proposal_content:
        content = st.session_state.proposal_content
        with st.expander("Preview del contenido generado", expanded=True):
            if content.introduction:
                st.markdown("**Introducción**")
                st.write(content.introduction)
            if content.company_presentation:
                st.markdown("**Presentación de la Empresa**")
                st.write(content.company_presentation)
            if content.technical_proposal:
                st.markdown("**Propuesta Técnica**")
                st.write(content.technical_proposal)
            if content.deliverables:
                st.markdown("**Entregables**")
                st.write(content.deliverables)
            if content.conclusion:
                st.markdown("**Conclusión**")
                st.write(content.conclusion)

    # Download button
    if st.session_state.generated_docx:
        safe_name = project_title.replace(" ", "_").replace("/", "-")[:50]
        filename = f"{st.session_state.ref_number}_{safe_name}.docx"

        st.download_button(
            label="Descargar .docx",
            data=st.session_state.generated_docx,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
            use_container_width=True,
        )
