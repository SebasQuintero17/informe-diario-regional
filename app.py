# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  INFORME DIARIO REGIONAL — Comfenalco Antioquia                             ║
# ║  100% gratuito: Groq (gratis) + Google News RSS + Streamlit Cloud           ║
# ╚══════════════════════════════════════════════════════════════════════════════╝

import streamlit as st
import feedparser
from groq import Groq
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime
import json, re, io, time, html, urllib.parse

# ── PAGE CONFIG ────────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Informe Diario Regional · Comfenalco",
    page_icon="🛡️",
    layout="wide",
)

# ══════════════════════════════════════════════════════════════════════════════
#  DATOS
# ══════════════════════════════════════════════════════════════════════════════
SUBREGIONES = {
    "BAJO CAUCA":         ["CACERES", "CAUCASIA", "EL BAGRE", "NECHI", "TARAZÁ", "ZARAGOZA"],
    "MAGDALENA MEDIO":    ["CARACOLI", "MACEO", "PUERTO BERRIO", "PUERTO NARE", "PUERTO TRIUNFO", "YONDO"],
    "NORDESTE":           ["AMALFI", "ANORI", "CISNEROS", "REMEDIOS", "SAN ROQUE", "SANTO DOMINGO", "SEGOVIA", "VEGACHI", "YALI", "YOLOMBO"],
    "NORTE":              ["ANGOSTURA", "BELMIRA", "BRICEÑO", "CAMPAMENTO", "CAROLINA DEL PRINCIPE", "DON MATIAS", "ENTRERRIOS", "GOMEZ PLATA", "GUADALUPE", "ITUANGO", "SAN ANDRES DE CUERQUIA", "SANTA ROSA DE OSOS", "TOLEDO", "VALDIVIA", "YARUMAL"],
    "OCCIDENTE":          ["ABRIAQUI", "ANZA", "BURITICA", "CAICEDO", "CAÑAS GORDAS", "DABEIBA", "EBEJICO", "FRONTINO", "HELICONIA", "LIBORINA", "OLAYA", "PEQUE", "SABANALARGA", "SAN JERONIMO", "SANTA FE DE ANTIOQUIA", "SOPETRÁN", "URAMITA"],
    "ORIENTE":            ["ABEJORRAL", "ALEJANDRIA", "ARGELIA", "COCORNÁ", "EL CARMEN DE VIBORAL", "EL PEÑOL", "EL RETIRO", "EL SANTUARIO", "GRANADA", "GUARNE", "GUATAPE", "LA CEJA", "LA UNION", "MARINILLA", "RIONEGRO", "SAN CARLOS", "SAN FRANCISCO", "SAN LUIS", "SAN RAFAEL", "SAN VICENTE", "SONSON"],
    "URABÁ":              ["APARTADÓ", "ARBOLETES", "CAREPA", "CHIGORODÓ", "MURINDO", "MUTATA", "NECOCLI", "SAN JUAN DE URABA", "SAN PEDRO DE URABA", "TURBO", "VIGIA DEL FUERTE"],
    "SUROESTE":           ["AMAGA", "ANDES", "ANGELOPOLIS", "BETANIA", "BETULIA", "CARAMANTA", "CIUDAD BOLIVAR", "CONCORDIA", "FREDONIA", "HISPANIA", "JARDIN", "JERICO", "LA PINTADA", "MONTEBELLO", "PUEBLORICO", "SALGAR", "SANTA BARBARA", "TAMESIS", "TARSO", "TITIRIBI", "URRAO", "VALPARAISO", "VENECIA"],
    "ÁREA METROPOLITANA": ["MEDELLÍN", "BELLO", "ENVIGADO", "ITAGÜÍ", "SABANETA", "LA ESTRELLA", "CALDAS", "COPACABANA", "GIRARDOTA", "BARBOSA"],
}

COLORES_EXCEL = {"GRAVES": "FF0000", "RELEVANTES": "FFC000", "PARCIALES": "FFFF00", "NINGUNA": "FFFFFF"}
LABELS_CRIT   = {"GRAVES": "AFECTACIONES GRAVES", "RELEVANTES": "AFECTACIONES RELEVANTES", "PARCIALES": "AFECTACIONES PARCIALES", "NINGUNA": "NO HAY AFECTACIONES"}
SEMAFORO      = {"GRAVES": "🔴", "RELEVANTES": "🟠", "PARCIALES": "🟡", "NINGUNA": "🟢"}

# ══════════════════════════════════════════════════════════════════════════════
#  CSS
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=IBM+Plex+Mono:wght@400;500&family=IBM+Plex+Sans:wght@300;400;600;700&display=swap');
html, body, [class*="css"] { font-family: 'IBM Plex Sans', sans-serif !important; }
.stApp { background-color: #070f0b !important; }
.block-container { padding-top: 1.2rem !important; max-width: 1280px !important; }
.hdr { background: linear-gradient(135deg,#0b1f14 0%,#122b1c 60%,#091610 100%);
       border:1px solid #1e5235; border-radius:14px; padding:1.8rem 2.4rem;
       margin-bottom:1.6rem; position:relative; overflow:hidden; }
.hdr::after { content:""; position:absolute; inset:0;
              background-image:repeating-linear-gradient(-45deg,transparent 0px,transparent 18px,rgba(52,168,83,.03) 18px,rgba(52,168,83,.03) 19px);
              pointer-events:none; }
.hdr-title { margin:0; font-size:1.9rem; font-weight:700; color:#52b788; letter-spacing:-.3px; }
.hdr-sub   { margin:4px 0 0; font-size:.75rem; font-family:'IBM Plex Mono',monospace; color:#74c69d; letter-spacing:1.2px; }
.hdr-badge { display:inline-block; margin-top:.8rem; background:#1b4332; color:#95d5b2;
             font-family:'IBM Plex Mono',monospace; font-size:.72rem;
             padding:3px 12px; border-radius:20px; border:1px solid #2d6a4f; }
.panel-title { font-size:.7rem; font-family:'IBM Plex Mono',monospace; color:#52b788; letter-spacing:1px; margin-bottom:.5rem; }
div[data-testid="stButton"] > button {
    background:linear-gradient(135deg,#1e5235,#2d7d4f) !important; color:#d8f3dc !important;
    border:1px solid #2d6a4f !important; border-radius:8px !important; font-size:.95rem !important;
    font-weight:600 !important; width:100% !important; padding:.75rem 1.4rem !important;
    box-shadow:0 4px 18px rgba(45,125,79,.35) !important; transition:all .2s !important; }
div[data-testid="stButton"] > button:hover { box-shadow:0 6px 24px rgba(45,125,79,.55) !important; transform:translateY(-1px) !important; }
div[data-testid="stDownloadButton"] > button { background:#111f16 !important; color:#74c69d !important; border:1px solid #2d6a4f !important; border-radius:8px !important; font-weight:600 !important; width:100% !important; }
div[data-testid="stProgress"] > div > div > div { background:linear-gradient(90deg,#2d6a4f,#52b788) !important; }
input { background:#0f1f16 !important; border-color:#1e3d2a !important; color:#c8ddd0 !important; border-radius:6px !important; }
details > summary { background:#0c1a12 !important; border:1px solid #1a3a26 !important; border-radius:8px !important; color:#c8ddd0 !important; }
div[data-testid="metric-container"] { background:#0c1a12; border:1px solid #182e1f; border-radius:8px; padding:.6rem 1rem; }
.placeholder { background:#0c1a12; border:2px dashed #1e4a30; border-radius:12px; padding:3rem 2rem; text-align:center; color:#3a6650; margin-top:.5rem; }
</style>
""", unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════════════════════════
#  BÚSQUEDA: Google News RSS — gratis, sin key
# ══════════════════════════════════════════════════════════════════════════════
def limpiar_html(texto: str) -> str:
    return html.unescape(re.sub(r'<[^<]+?>', '', texto or '')).strip()


def buscar_noticias(query: str, max_resultados: int = 8) -> list[dict]:
    try:
        url = (
            "https://news.google.com/rss/search"
            f"?q={urllib.parse.quote(query)}&hl=es&gl=CO&ceid=CO:es"
        )
        feed = feedparser.parse(url)
        return [
            {
                "titulo":  limpiar_html(e.get("title", "")),
                "resumen": limpiar_html(e.get("summary", ""))[:400],
                "fuente":  e.get("source", {}).get("title", ""),
            }
            for e in feed.entries[:max_resultados]
        ]
    except Exception:
        return []


def formatear(resultados: list[dict]) -> str:
    if not resultados:
        return "(Sin resultados de búsqueda)"
    return "\n".join(
        f"• [{r['fuente']}] {r['titulo']}: {r['resumen']}"
        for r in resultados
    )


def obtener_noticias(fecha: str) -> tuple[str, str, str]:
    queries = [
        f"Antioquia Colombia seguridad orden público {fecha}",
        f"vías carreteras Antioquia Colombia cierre derrumbe {fecha}",
        f"Antioquia Colombia servicios públicos agua energía comercio {fecha}",
    ]
    resultados = []
    for i, q in enumerate(queries):
        if i > 0:
            time.sleep(1)
        resultados.append(formatear(buscar_noticias(q)))
    return resultados[0], resultados[1], resultados[2]


# ══════════════════════════════════════════════════════════════════════════════
#  IA: Groq Llama 3.3 70B — gratis, sin tarjeta
# ══════════════════════════════════════════════════════════════════════════════
def construir_prompt(seg: str, vias: str, svc: str, fecha: str) -> str:
    lista = "\n".join(f"  • {s}: {', '.join(m[:5])}" for s, m in SUBREGIONES.items())
    return f"""Eres analista de seguridad y movilidad de Comfenalco Antioquia. Fecha: {fecha}.

SUBREGIONES DE ANTIOQUIA:
{lista}

─── NOTICIAS SEGURIDAD / ORDEN PÚBLICO ───
{seg}

─── NOTICIAS VÍAS Y MOVILIDAD ───
{vias}

─── NOTICIAS SERVICIOS PÚBLICOS Y COMERCIO ───
{svc}

INSTRUCCIONES:
• Asigna cada noticia a la subregión del municipio mencionado.
• Sin información para una categoría → "SIN NOVEDAD".
• Criticidad: GRAVES=toques de queda/desplazamientos/cierres totales,
  RELEVANTES=alertas/ataques/pasos a un carril,
  PARCIALES=seguimientos/obras/cortes programados,
  NINGUNA=sin novedad.

Responde SOLO con JSON válido, sin texto adicional:
{{
  "BAJO CAUCA":         {{"municipio_principal":"","orden_publico":"","orden_publico_crit":"","vias":"","vias_crit":"","comercio_servicios":"","comercio_crit":"","sedes_hoteles":"","sedes_crit":"","criticidad_general":""}},
  "MAGDALENA MEDIO":    {{"municipio_principal":"","orden_publico":"","orden_publico_crit":"","vias":"","vias_crit":"","comercio_servicios":"","comercio_crit":"","sedes_hoteles":"","sedes_crit":"","criticidad_general":""}},
  "NORDESTE":           {{"municipio_principal":"","orden_publico":"","orden_publico_crit":"","vias":"","vias_crit":"","comercio_servicios":"","comercio_crit":"","sedes_hoteles":"","sedes_crit":"","criticidad_general":""}},
  "NORTE":              {{"municipio_principal":"","orden_publico":"","orden_publico_crit":"","vias":"","vias_crit":"","comercio_servicios":"","comercio_crit":"","sedes_hoteles":"","sedes_crit":"","criticidad_general":""}},
  "OCCIDENTE":          {{"municipio_principal":"","orden_publico":"","orden_publico_crit":"","vias":"","vias_crit":"","comercio_servicios":"","comercio_crit":"","sedes_hoteles":"","sedes_crit":"","criticidad_general":""}},
  "ORIENTE":            {{"municipio_principal":"","orden_publico":"","orden_publico_crit":"","vias":"","vias_crit":"","comercio_servicios":"","comercio_crit":"","sedes_hoteles":"","sedes_crit":"","criticidad_general":""}},
  "URABÁ":              {{"municipio_principal":"","orden_publico":"","orden_publico_crit":"","vias":"","vias_crit":"","comercio_servicios":"","comercio_crit":"","sedes_hoteles":"","sedes_crit":"","criticidad_general":""}},
  "SUROESTE":           {{"municipio_principal":"","orden_publico":"","orden_publico_crit":"","vias":"","vias_crit":"","comercio_servicios":"","comercio_crit":"","sedes_hoteles":"","sedes_crit":"","criticidad_general":""}},
  "ÁREA METROPOLITANA": {{"municipio_principal":"","orden_publico":"","orden_publico_crit":"","vias":"","vias_crit":"","comercio_servicios":"","comercio_crit":"","sedes_hoteles":"","sedes_crit":"","criticidad_general":""}}
}}"""


def analizar(api_key: str, seg: str, vias: str, svc: str, fecha: str) -> dict:
    client = Groq(api_key=api_key)
    try:
        resp = client.chat.completions.create(
            model="llama-3.3-70b-versatile",
            messages=[{"role": "user", "content": construir_prompt(seg, vias, svc, fecha)}],
            max_tokens=4096,
            temperature=0.1,
        )
        raw = resp.choices[0].message.content
        match = re.search(r'\{[\s\S]*\}', raw)
        if match:
            return json.loads(match.group())
    except Exception as e:
        st.warning(f"⚠️ Error Groq: {e}")

    return {
        sub: {"municipio_principal": m[0],
              "orden_publico": "SIN NOVEDAD", "orden_publico_crit": "NINGUNA",
              "vias": "SIN NOVEDAD",           "vias_crit": "NINGUNA",
              "comercio_servicios": "SIN NOVEDAD", "comercio_crit": "NINGUNA",
              "sedes_hoteles": "SIN NOVEDAD",  "sedes_crit": "NINGUNA",
              "criticidad_general": "NINGUNA"}
        for sub, m in SUBREGIONES.items()
    }


# ══════════════════════════════════════════════════════════════════════════════
#  EXCEL
# ══════════════════════════════════════════════════════════════════════════════
def construir_excel(datos: dict, fecha: str) -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "REPORTE"

    def fill(h): return PatternFill("solid", fgColor=h)
    def borde():
        s = Side(style="thin", color="A8C4B0")
        return Border(left=s, right=s, top=s, bottom=s)
    def hdr(c, bg="2D6A4F"):
        c.font = Font(bold=True, color="FFFFFF", size=9)
        c.fill = fill(bg); c.border = borde()
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    def dato(c, bg="FFFFFF"):
        c.fill = fill(bg[-6:]); c.border = borde()
        c.alignment = Alignment(wrap_text=True, vertical="top", horizontal="left")
        c.font = Font(size=9)

    # Título
    ws.merge_cells("A1:G1")
    t = ws["A1"]
    t.value = "INFORME DIARIO REGIONAL — COMFENALCO ANTIOQUIA"
    t.font = Font(bold=True, size=13, color="FFFFFF"); t.fill = fill("0D2818")
    t.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 32

    # Fecha
    ws.merge_cells("A2:G2")
    d = ws["A2"]
    d.value = f"Fecha: {fecha}     |     Generado automáticamente con Inteligencia Artificial"
    d.font = Font(size=9, italic=True, color="5A8A6A")
    d.alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 18

    # Encabezados
    for col, h in enumerate([
        "SUBREGIÓN", "MUNICIPIOS",
        "ORDEN PÚBLICO\n(Acciones de grupos criminales)",
        "AFECTACIÓN DE VÍAS\n(Eventos que afecten vías)",
        "ESTADO COMERCIO O SERVICIOS\n(Alteraciones económicas)",
        "ESTADO DE SEDES, HOTELES Y PARQUES\n(Afectación de servicios públicos)",
        "CRITICIDAD",
    ], 1):
        hdr(ws.cell(row=3, column=col, value=h))
    ws.row_dimensions[3].height = 44

    # Datos
    row = 4
    for sub, d in datos.items():
        bg = lambda k: COLORES_EXCEL.get(d.get(k, "NINGUNA"), "FFFFFF")
        cg = d.get("criticidad_general", "NINGUNA")
        filas = [
            (1, sub,                                       bg("criticidad_general")),
            (2, d.get("municipio_principal", ""),          bg("criticidad_general")),
            (3, d.get("orden_publico",  "SIN NOVEDAD"),    bg("orden_publico_crit")),
            (4, d.get("vias",           "SIN NOVEDAD"),    bg("vias_crit")),
            (5, d.get("comercio_servicios","SIN NOVEDAD"), bg("comercio_crit")),
            (6, d.get("sedes_hoteles",  "SIN NOVEDAD"),    bg("sedes_crit")),
            (7, LABELS_CRIT.get(cg, ""),                   bg("criticidad_general")),
        ]
        for col, val, color in filas:
            c = ws.cell(row=row, column=col, value=val)
            dato(c, color)
            if col in (1, 2, 7):
                c.font = Font(bold=True, size=9)
                c.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
        ws.row_dimensions[row].height = 75
        row += 1

    # Leyenda
    row += 1
    ws.merge_cells(f"A{row}:B{row}")
    ws[f"A{row}"].value = "NIVEL CRITICIDAD ORDEN PÚBLICO A LA FECHA"
    ws[f"A{row}"].font = Font(bold=True, size=9)
    for i, (label, color, nivel) in enumerate([
        ("NO HAY\nAFECTACIONES",   "92D050", "BAJA"),
        ("AFECTACIONES\nPARCIALES", "FFFF00", "MEDIA"),
        ("AFECTACIONES\nRELEVANTES","FFC000", "ALTA"),
        ("AFECTACIONES\nGRAVES",    "FF0000", "MUY ALTA"),
    ]):
        r = row + 1 + i
        c = ws.cell(row=r, column=1, value=label)
        c.fill = fill(color); c.font = Font(bold=True, size=9); c.border = borde()
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
        ws.cell(row=r, column=3, value=nivel).font = Font(bold=True, size=9)
        ws.row_dimensions[r].height = 28

    # Contactos
    rc = row + 1
    for col, txt in [
        (4, "INVÍAS / X: @numeral767\nMintransporte / X: @MintransporteCo\nPolicía Vías / X: @PoliciaDeTransito"),
        (5, "INVÍAS WhatsApp: 3176460735 / Línea: 01 8000 112 137\nMintransporte: 601 307 7733"),
        (6, "invias.gov.co\ninvias-viajero.vercel.app\npolicia.gov.co/estado-de-las-vias"),
        (7, "CIMCA: 3104382572\nINVÍAS: 018005191656\nEn ruta WhatsApp: 317 646 0735\nEmergencias: 123"),
    ]:
        c = ws.cell(row=rc, column=col, value=txt)
        c.font = Font(size=8)
        c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[rc].height = 60

    # Anchos
    for i, w in enumerate([20, 20, 44, 44, 38, 38, 22], 1):
        ws.column_dimensions[get_column_letter(i)].width = w

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf


# ══════════════════════════════════════════════════════════════════════════════
#  UI
# ══════════════════════════════════════════════════════════════════════════════
hoy = datetime.now()

st.markdown(f"""
<div class="hdr">
  <div style="display:flex;align-items:center;gap:1.2rem;position:relative;z-index:1">
    <span style="font-size:2.6rem">🛡️</span>
    <div>
      <h1 class="hdr-title">Informe Diario Regional</h1>
      <p class="hdr-sub">COMFENALCO ANTIOQUIA &nbsp;·&nbsp; ÁREA DE SEGURIDAD &nbsp;·&nbsp; GENERACIÓN AUTOMÁTICA 100% GRATIS</p>
      <span class="hdr-badge">{hoy.strftime("%A %d de %B de %Y").upper()}</span>
    </div>
  </div>
</div>
""", unsafe_allow_html=True)

izq, der = st.columns([1, 2.3], gap="large")

with izq:
    st.markdown('<p class="panel-title">⚙️ CONFIGURACIÓN</p>', unsafe_allow_html=True)

    try:
        groq_key = st.secrets["GROQ_API_KEY"]
        st.success("✅ Groq API Key configurada")
    except Exception:
        groq_key = st.text_input(
            "🔑 Groq API Key (gratis)",
            type="password",
            placeholder="gsk_...",
            help="Gratis en console.groq.com — sin tarjeta de crédito",
        )

    fecha_informe = st.date_input("📅 Fecha del informe", value=hoy)

    st.markdown("---")
    st.markdown('<p class="panel-title">📍 SUBREGIONES</p>', unsafe_allow_html=True)
    seleccion = {sub: st.checkbox(sub, value=True, key=f"chk_{sub}") for sub in SUBREGIONES}

    n_activas = sum(seleccion.values())
    st.markdown("---")
    st.markdown(f"""
    <div style="font-size:.72rem;color:#3d7054;font-family:'IBM Plex Mono',monospace;line-height:2">
    📡 Subregiones: <b style="color:#74c69d">{n_activas}</b><br>
    ⏱️ Tiempo estimado: ~30–60 seg<br>
    🔍 Fuente: Google News RSS<br>
    🤖 IA: Groq · Llama 3.3 70B<br>
    💰 Costo: <b style="color:#52b788">$0.00</b>
    </div>
    """, unsafe_allow_html=True)

with der:
    st.markdown('<p class="panel-title">🚀 GENERAR INFORME</p>', unsafe_allow_html=True)
    generar = st.button("⚡ GENERAR INFORME AUTOMÁTICO", use_container_width=True)

    if generar:
        if not groq_key:
            st.error("❌ Necesitas la Groq API Key. Consíguela gratis en console.groq.com")
            st.stop()

        fecha_str = fecha_informe.strftime("%d/%m/%Y")
        barra  = st.progress(0)
        estado = st.empty()

        # Paso 1: Noticias
        estado.markdown("🔍 **Buscando noticias en Google News...** (seguridad · vías · servicios)")
        barra.progress(10)
        n_seg, n_vias, n_svc = obtener_noticias(fecha_str)
        barra.progress(40)

        # Paso 2: IA
        estado.markdown("🤖 **Analizando con IA...** Groq · Llama 3.3 70B")
        datos = analizar(groq_key, n_seg, n_vias, n_svc, fecha_str)
        datos = {s: datos[s] for s in SUBREGIONES if seleccion.get(s) and s in datos}
        barra.progress(85)

        # Paso 3: Excel
        estado.markdown("📊 **Generando Excel...**")
        excel_buf = construir_excel(datos, fecha_str)
        barra.progress(100)
        estado.markdown("✅ **Informe listo**")

        # Métricas
        st.markdown("---")
        conteo = {"GRAVES": 0, "RELEVANTES": 0, "PARCIALES": 0, "NINGUNA": 0}
        for d in datos.values():
            c = d.get("criticidad_general", "NINGUNA")
            conteo[c] = conteo.get(c, 0) + 1

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("🔴 Graves",      conteo["GRAVES"])
        m2.metric("🟠 Relevantes",  conteo["RELEVANTES"])
        m3.metric("🟡 Parciales",   conteo["PARCIALES"])
        m4.metric("🟢 Sin novedad", conteo["NINGUNA"])

        # Vista previa
        st.markdown("#### 📋 Resumen de Novedades")
        for sub, d in datos.items():
            crit  = d.get("criticidad_general", "NINGUNA")
            emoji = SEMAFORO.get(crit, "⚪")
            mun   = d.get("municipio_principal", "")
            label = LABELS_CRIT.get(crit, "")
            with st.expander(f"{emoji}  {sub}  —  {mun}  ·  {label}"):
                c1, c2 = st.columns(2)
                with c1:
                    st.markdown(f"**🚨 Orden Público** {SEMAFORO.get(d.get('orden_publico_crit','NINGUNA'),'')}")
                    st.caption(d.get("orden_publico", "SIN NOVEDAD"))
                    st.markdown(f"**🛣️ Vías** {SEMAFORO.get(d.get('vias_crit','NINGUNA'),'')}")
                    st.caption(d.get("vias", "SIN NOVEDAD"))
                with c2:
                    st.markdown(f"**🏪 Comercio** {SEMAFORO.get(d.get('comercio_crit','NINGUNA'),'')}")
                    st.caption(d.get("comercio_servicios", "SIN NOVEDAD"))
                    st.markdown(f"**🏢 Sedes / Hoteles** {SEMAFORO.get(d.get('sedes_crit','NINGUNA'),'')}")
                    st.caption(d.get("sedes_hoteles", "SIN NOVEDAD"))

        # Descarga
        st.markdown("---")
        nombre = f"Informe_Diario_Regional_{fecha_informe.strftime('%Y%m%d')}.xlsx"
        st.download_button(
            label="📥 DESCARGAR EXCEL",
            data=excel_buf, file_name=nombre,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.caption(f"Listo para revisar y convertir a PDF: `{nombre}`")

    else:
        st.markdown("""
        <div class="placeholder">
          <div style="font-size:3rem;margin-bottom:.8rem">📡</div>
          <div style="font-size:1rem;font-weight:600;color:#52b788">Sistema listo</div>
          <div style="font-size:.82rem;margin-top:.5rem;line-height:1.8;color:#3d7054">
            Configura la fecha a la izquierda<br>
            y presiona el botón para generar.
          </div>
        </div>
        """, unsafe_allow_html=True)
