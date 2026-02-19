# El Licitador - Product Vision

## Resumen Ejecutivo

**Cliente:** Jorgito (Director Comercial, Educabot)
**Inversión:** USD 8.000 (implementación) + 3 reuniones de iteración
**Timeline:** 3 semanas (1 reunión/semana)
**Fecha:** 2026-02-18

---

## 1. El Problema

Jorgito factura ~400K/año con ~200K de margen bruto gestionando el área comercial de Educabot. Genera dos tipos de documentos constantemente:

### Respuestas a Licitaciones
- Recibe un pliego y debe responder **punto por punto** confirmando cumplimiento
- Debe incluir **experiencia vinculante**: proyectos previos relacionados con lo que se pide
- El contenido es repetitivo: ya tiene la info, ya usa ChatGPT para redactar
- **El dolor real: el formato.** Horas copiando, pegando, formateando Word/Google Docs

### Propuestas Comerciales / Presupuestos
- Estructura variable según contexto (vendido vs. no vendido, formal vs. informal)
- Incluyen: intro, propuesta técnica (a veces), tabla de precios (Excel), condiciones, impacto esperado
- La tabla de precios tiene colores corporativos, tipografía, fórmulas
- El número de referencia lo inventa cada vez (sin sistema)

### El dolor cuantificado
- **Formato** = horas de trabajo manual en CADA documento
- Encabezados, pie de página (firma de Felipe), numeración real de Google Docs, estilos, colores
- Copy-paste entre documentos rompe el formato → rehacer manualmente
- Es "como hacer un PowerPoint dentro de un Word"
- El contenido NO es el dolor (ChatGPT ya lo resuelve aceptablemente)

---

## 2. Propuesta de Valor

> **Eliminar las horas de formateo** generando documentos comerciales listos para enviar, con el contenido correcto, el formato correcto, y la experiencia vinculante correcta, en minutos en vez de horas.

### Para quién
Jorgito como usuario principal. Él es quien tiene las "reglas de Felipe" internalizadas y el criterio comercial. Su equipo NO es usuario inicial.

### El cambio esperado
| Antes | Después |
|-------|---------|
| Horas formateando cada documento | Documento formateado listo en minutos |
| Inventar N° de referencia cada vez | Sistema automático de referenciación |
| Buscar licitaciones anteriores manualmente | Base de conocimiento con búsqueda inteligente |
| Copy-paste que rompe formato | Output nativo en Word/Google Docs con estilos |
| Tabla de precios hecha a mano | Tabla generada con formato corporativo |

---

## 3. Alcance del Producto (MVP)

### 3.1 Flujo: Respuesta a Licitación

```
1. Jorgito sube/pega el pliego de licitación
2. El sistema lo analiza y extrae los puntos a responder
3. Le pregunta: "¿Qué productos vas a ofrecer?" (o lo infiere del pliego)
4. Busca en la base de conocimiento licitaciones anteriores relacionadas
5. Presenta: "Encontré estas 5 experiencias vinculantes. ¿Cuáles incluimos?"
6. Genera la respuesta completa punto por punto
7. Exporta como Word (.docx) con formato correcto:
   - Encabezado corporativo
   - Pie de página con firma de Felipe
   - Numeración real (no texto plano)
   - Estilos y tipografía correctos
```

**Caso especial: Licitaciones CABA**
- Formato específico de presentación
- Consignas predefinidas que siempre se incluyen
- Referencias a proyectos anteriores con CABA
- Template dedicado

### 3.2 Flujo: Propuesta Comercial / Presupuesto

```
1. Jorgito indica: cliente, productos, contexto (vendido/no vendido, formal/informal)
2. El sistema determina el nivel de detalle:
   - Vendido/informal → Solo tabla de precios + condiciones comerciales
   - No vendido/formal → Documento completo con descripción técnica
3. Genera el documento con la estructura correcta:
   - Fecha
   - N° de referencia (generado con criterio consistente)
   - Introducción
   - Propuesta técnica (si aplica)
   - Propuesta comercial (tabla con formato corporativo)
   - Condiciones comerciales
   - Impacto esperado (si aplica)
   - Conclusión (si aplica)
   - Anexos (si aplica)
4. La tabla de precios se genera como tabla formateada con colores corporativos
5. Exporta como Word (.docx) con formato completo
```

### 3.3 Base de Conocimiento

- **Productos Educabot**: Descripciones de CODI, robots, plataformas, train, AI assistants
- **Licitaciones anteriores**: Indexadas por tipo, cliente, productos ofrecidos
- **Propuestas anteriores**: Templates y ejemplos de propuestas de Martín, hardware, software
- **Experiencia vinculante**: Proyectos previos categorizados por tipo de solución
- **Info CABA**: Formato específico, consignas, historial de presentaciones
- **Contratos tipo**: Estructura de contratos (ej: Taro)

### 3.4 Sistema de Referencias

Formato propuesto: `EDU-{TIPO}-{AÑO}-{SECUENCIAL}`

Ejemplos:
- `EDU-LIC-26-001` → Licitación #1 del 2026
- `EDU-PRO-26-015` → Propuesta comercial #15 del 2026
- `EDU-PRO-26-015-R2` → Revisión 2 de la propuesta #15

---

## 4. Approach Técnico

### Stack: Streamlit + python-docx + Claude Agent SDK

**Arquitectura:**
```
┌──────────────────────────────────┐
│  Browser (Streamlit UI)          │
│  ├─ Upload de pliego/contexto    │
│  ├─ Chat interactivo             │
│  ├─ Preview del documento        │
│  └─ Botón de descarga .docx     │
└──────────┬───────────────────────┘
           │
┌──────────▼───────────────────────┐
│  Python Backend                  │
│  ├─ Claude Agent SDK (LLM)      │
│  ├─ python-docx (generación)    │
│  ├─ Base de conocimiento        │
│  └─ Sistema de referencias      │
└──────────────────────────────────┘
```

**Por qué este stack:**

| Componente | Decisión | Razón |
|------------|----------|-------|
| **UI** | Streamlit | Browser-based, Jorgito no toca terminal. Upload, chat, preview y download out-of-the-box. 1-2 días de build. |
| **LLM** | Claude Agent SDK (Python) | Wraps Claude Code como subprocess. Acceso a contexto, skills, y herramientas. |
| **Generador DOCX** | `python-docx` | Mismo approach que usa Claude Desktop (claude.ai). Librería madura, control total de estilos, tablas, headers/footers. |
| **Referencia** | Skill `anthropics/skills@docx` instalada | 12K+ installs, documenta gotchas críticos (DXA units, ShadingType.CLEAR, dual widths en tablas). Usamos sus patterns como guía. |

**Nota sobre la skill de Anthropic:**
- La skill oficial usa `docx-js` (JavaScript) porque Claude Code opera en Node
- Claude Desktop (claude.ai) usa `python-docx` en su sandbox
- Nosotros usamos `python-docx` (match con Streamlit/Python) pero aplicamos los mismos patterns y gotchas documentados en la skill

**Componentes a desarrollar:**
1. **Streamlit App** → UI completa con upload, chat, preview, download
2. **Flujo Licitación** → Lógica de análisis de pliego + generación de respuesta
3. **Flujo Propuesta** → Lógica de propuestas comerciales con tabla de precios
4. **Template DOCX** → Template python-docx con estilos corporativos de Educabot (encabezado, pie con firma Felipe, numeración, colores)
5. **Base de conocimiento** → Productos, licitaciones anteriores, experiencia vinculante
6. **Sistema de referencias** → Generación automática de N° de referencia

### Alternativas consideradas

| Opción | Pros | Contras | Decisión |
|--------|------|---------|----------|
| **Streamlit + python-docx** | UI amigable, todo Python, rápido de construir | Hosting necesario (local o cloud) | **Elegido** |
| **Claude Code CLI + Skills** | Simple, sin UI | Jorgito no es técnico, terminal intimidante | Descartado |
| **Tauri/Electron** | App nativa, UX pulida | 1-2 semanas de build, over-engineering para MVP | Descartado (v2 si necesario) |
| **GPT personalizado** | Ya lo conoce | No genera DOCX formateado con estilos reales | Descartado |
| **CUI / Claudia (wrappers existentes)** | Ya construidos | No son document-focused, habría que customizar igual | Descartado |

---

## 5. Lo que NO incluye el MVP

- Integración con Google Docs API (el output es .docx que se sube manual)
- Gestión de clientes / CRM
- Capacitación al equipo de Jorgito (fuera de scope, venta futura)
- Integración con Excel de propuestas existente (el sistema genera su propia tabla)
- Hosting en la nube (v1 corre local en la máquina de Jorgito)

---

## 6. Plan de Entrega

### Semana 1: Setup + Licitaciones
- Recopilar: 3-5 licitaciones anteriores, productos Educabot, formato CABA
- Armar base de conocimiento
- Crear Skill de Licitaciones (flujo completo)
- Crear generador DOCX con template corporativo
- **Reunión 1:** Demo + feedback + ajustes

### Semana 2: Propuestas + Refinamiento
- Recopilar: 3-5 propuestas comerciales anteriores, template de Martín
- Crear Skill de Propuestas Comerciales
- Implementar sistema de N° de referencia
- Refinar formato DOCX basado en feedback de Semana 1
- **Reunión 2:** Demo + feedback + ajustes

### Semana 3: Pulido + Entrega
- Ajustes finales de formato y contenido
- Documentación de uso (capacitación grabada)
- Entrega de todo el sistema
- **Reunión 3:** Entrega final + cierre

---

## 7. Métricas de Éxito

| Métrica | Objetivo |
|---------|----------|
| Tiempo de creación de licitación | De horas → <30 min |
| Tiempo de creación de propuesta | De horas → <15 min |
| Tiempo de formateo manual post-generación | <10 min de ajustes |
| Adopción | Jorgito lo usa para el 80%+ de sus documentos al mes |

---

## 8. Riesgos y Supuestos

### Supuestos
- Jorgito se siente cómodo usando Claude Code (o una interfaz CLI simplificada)
- Las "reglas de Felipe" se pueden codificar en instrucciones claras
- El formato DOCX generado programáticamente es suficientemente bueno vs. Google Docs nativo
- 3-5 documentos de ejemplo son suficientes para capturar los patrones

### Riesgos
| Riesgo | Mitigación |
|--------|-----------|
| Jorgito no se adapta a CLI | Crear aliases simples (`licitacion`, `propuesta`) o wrapper con UI mínima |
| Formato DOCX no coincide exactamente | Semana 1 se dedica a calibrar el template con feedback directo |
| Base de conocimiento insuficiente | Iterar el contenido durante las 3 reuniones |
| Cada licitación es demasiado diferente | El flujo es interactivo, no automático: siempre confirma con Jorgito |

---

## 9. Oportunidades Futuras (Post-MVP)

- **Venta al equipo**: Capacitación + licencias individuales (nueva venta)
- **Integración Google Docs API**: Generar directamente en Drive con formato nativo
- **Dashboard de propuestas**: Tracking de propuestas enviadas, estado, win rate
- **Integración con Excel**: Vincular tabla de precios con spreadsheet maestro
- **Templates por vertical**: Educación, gobierno, corporativo, hardware, software
