import streamlit as st
import pandas as pd
import os
from datetime import datetime, date, timedelta
from io import BytesIO
import base64

st.set_page_config(
    page_title="PetShop Manager Pro - Demo",
    page_icon="🐾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ======================================================
# MODO DEMO BLOQUEADO
# ======================================================
# En True: el cliente puede mirar la app, navegar y ver datos,
# pero NO puede guardar ventas, compras, stock, clientes, chat,
# pérdidas, importar datos ni modificar información.
DEMO_BLOQUEADA = True

DATA_DIR = "data_petshop"
CHAT_DIR = f"{DATA_DIR}/chat_adjuntos"
os.makedirs(DATA_DIR, exist_ok=True)
os.makedirs(CHAT_DIR, exist_ok=True)

VENTAS_FILE = f"{DATA_DIR}/ventas.csv"
STOCK_FILE = f"{DATA_DIR}/stock.csv"
COMPRAS_FILE = f"{DATA_DIR}/compras.csv"
CLIENTES_FILE = f"{DATA_DIR}/clientes.csv"
PERDIDAS_FILE = f"{DATA_DIR}/perdidas.csv"
CHAT_FILE = f"{DATA_DIR}/chat.csv"
LOGO_FILE = "logo.png"

SUCURSALES = ["Casa Central", "Sucursal Norte", "Sucursal Sur"]
CATEGORIAS = [
    "Alimento para perros", "Alimento para gatos", "Snacks", "Piedras sanitarias",
    "Arena para gatos", "Juguetes", "Correas", "Collares", "Camas", "Comederos",
    "Bebederos", "Shampoo", "Antipulgas", "Medicamentos", "Accesorios",
    "Ropa para mascotas", "Productos de higiene", "Otros"
]
MEDIOS_PAGO = ["Efectivo", "Débito", "Crédito", "Transferencia", "Mercado Pago"]

USUARIOS = {
    "demo": {"password": "1234", "rol": "Demo", "sucursal": "Todas", "nombre": "Cliente Demo"},
    "admin": {"password": "1234", "rol": "Administrador", "sucursal": "Todas", "nombre": "Dueño"},
    "encargado": {"password": "1234", "rol": "Encargado", "sucursal": "Casa Central", "nombre": "Encargado Casa Central"},
    "vendedor": {"password": "1234", "rol": "Vendedor", "sucursal": "Casa Central", "nombre": "Vendedor Casa Central"},
    "norte": {"password": "1234", "rol": "Vendedor", "sucursal": "Sucursal Norte", "nombre": "Vendedor Norte"},
    "sur": {"password": "1234", "rol": "Vendedor", "sucursal": "Sucursal Sur", "nombre": "Vendedor Sur"},
}

# ======================================================
# FUNCIONES
# ======================================================
def cargar_logo_base64(path):
    if os.path.exists(path):
        with open(path, "rb") as img:
            return base64.b64encode(img.read()).decode()
    return None


def pesos(valor):
    try:
        return f"$ {float(valor):,.0f}".replace(",", ".")
    except Exception:
        return "$ 0"


def cargar_csv(path):
    if os.path.exists(path):
        return pd.read_csv(path)
    return pd.DataFrame()


def guardar_csv(df, path):
    if DEMO_BLOQUEADA:
        st.warning("🔒 Modo demo bloqueado: esta acción está deshabilitada para clientes.")
        return False
    df.to_csv(path, index=False)
    return True


def guardar_csv_sistema(df, path):
    df.to_csv(path, index=False)


def exportar_excel(dfs):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for nombre, df in dfs.items():
            df.to_excel(writer, index=False, sheet_name=nombre[:31])
    return output.getvalue()


def metric_card(titulo, valor, icono, detalle=""):
    st.markdown(
        f"""
        <div class="metric-card">
            <div class="metric-top">
                <span>{titulo}</span>
                <span class="metric-icon">{icono}</span>
            </div>
            <div class="metric-value">{valor}</div>
            <div class="metric-detail">{detalle}</div>
        </div>
        """,
        unsafe_allow_html=True
    )


def crear_archivos_iniciales():
    hoy = date.today()
    if not os.path.exists(STOCK_FILE):
        stock = pd.DataFrame([
            {"codigo":"DOG001","producto":"Royal Canin Mini Adult 3kg","categoria":"Alimento para perros","marca":"Royal Canin","proveedor":"Distribuidora Mascotas","sucursal":"Casa Central","precio_compra":15000,"precio_venta":23000,"stock":12,"stock_minimo":5,"fecha_ingreso":str(hoy),"fecha_vencimiento":""},
            {"codigo":"DOG002","producto":"Alimento Perro Adulto 15kg","categoria":"Alimento para perros","marca":"DogPlus","proveedor":"Proveedor Norte","sucursal":"Sucursal Norte","precio_compra":18500,"precio_venta":28500,"stock":7,"stock_minimo":6,"fecha_ingreso":str(hoy),"fecha_vencimiento":""},
            {"codigo":"CAT001","producto":"Piedras Sanitarias 4kg","categoria":"Piedras sanitarias","marca":"MichiClean","proveedor":"Proveedor B","sucursal":"Sucursal Norte","precio_compra":2500,"precio_venta":4300,"stock":25,"stock_minimo":8,"fecha_ingreso":str(hoy),"fecha_vencimiento":""},
            {"codigo":"ACC001","producto":"Correa Retráctil Talla M","categoria":"Correas","marca":"PetFlex","proveedor":"Proveedor C","sucursal":"Sucursal Sur","precio_compra":6000,"precio_venta":11500,"stock":4,"stock_minimo":5,"fecha_ingreso":str(hoy),"fecha_vencimiento":""},
            {"codigo":"SNK001","producto":"Snacks Dentales Perro","categoria":"Snacks","marca":"HappyPet","proveedor":"Proveedor A","sucursal":"Casa Central","precio_compra":1800,"precio_venta":3500,"stock":31,"stock_minimo":10,"fecha_ingreso":str(hoy),"fecha_vencimiento":""},
        ])
        guardar_csv_sistema(stock, STOCK_FILE)

    if not os.path.exists(VENTAS_FILE):
        ventas_demo = []
        productos = [
            ("Casa Central", "DOG001", "Royal Canin Mini Adult 3kg", "Alimento para perros", 23000, 15000),
            ("Sucursal Norte", "CAT001", "Piedras Sanitarias 4kg", "Piedras sanitarias", 4300, 2500),
            ("Sucursal Sur", "ACC001", "Correa Retráctil Talla M", "Correas", 11500, 6000),
            ("Casa Central", "SNK001", "Snacks Dentales Perro", "Snacks", 3500, 1800),
        ]
        for i in range(24):
            suc, cod, prod, cat, precio, costo = productos[i % len(productos)]
            cant = (i % 4) + 1
            fecha = datetime.now() - timedelta(days=23 - i)
            total = precio * cant
            ventas_demo.append({
                "fecha": str(fecha), "sucursal": suc, "codigo": cod, "producto": prod,
                "categoria": cat, "cantidad": cant, "precio_unitario": precio,
                "descuento": 0, "total": total, "costo_total": costo * cant,
                "ganancia": total - costo * cant, "medio_pago": MEDIOS_PAGO[i % len(MEDIOS_PAGO)],
                "vendedor": "demo", "cliente": f"Cliente {i+1}", "comprobante": f"D-{1000+i}"
            })
        guardar_csv_sistema(pd.DataFrame(ventas_demo), VENTAS_FILE)

    if not os.path.exists(COMPRAS_FILE):
        compras = pd.DataFrame([
            {"fecha":str(datetime.now()),"proveedor":"Distribuidora Mascotas","codigo":"DOG001","producto":"Royal Canin Mini Adult 3kg","categoria":"Alimento para perros","sucursal":"Casa Central","cantidad":20,"precio_compra_unitario":15000,"costo_total":300000},
            {"fecha":str(datetime.now()),"proveedor":"Proveedor B","codigo":"CAT001","producto":"Piedras Sanitarias 4kg","categoria":"Piedras sanitarias","sucursal":"Sucursal Norte","cantidad":30,"precio_compra_unitario":2500,"costo_total":75000},
        ])
        guardar_csv_sistema(compras, COMPRAS_FILE)

    if not os.path.exists(CLIENTES_FILE):
        clientes = pd.DataFrame([
            {"nombre":"María López","telefono":"11-5555-1111","mascota":"Toby","tipo_mascota":"Perro","email":"maria@email.com","observaciones":"Compra alimento mensual"},
            {"nombre":"Juan Pérez","telefono":"11-5555-2222","mascota":"Milo","tipo_mascota":"Gato","email":"juan@email.com","observaciones":"Cliente frecuente"},
        ])
        guardar_csv_sistema(clientes, CLIENTES_FILE)

    if not os.path.exists(PERDIDAS_FILE):
        perdidas = pd.DataFrame([
            {"fecha":str(datetime.now()),"sucursal":"Sucursal Sur","codigo":"ACC001","producto":"Correa Retráctil Talla M","motivo":"Rotura","cantidad":1,"costo_unitario":6000,"perdida_total":6000,"observaciones":"Producto dañado en exhibición"}
        ])
        guardar_csv_sistema(perdidas, PERDIDAS_FILE)

    if not os.path.exists(CHAT_FILE):
        chat = pd.DataFrame([
            {"fecha":str(datetime.now()),"tipo_chat":"Chat general","canal":"general","de_usuario":"admin","de_nombre":"Dueño","para_usuario":"","mensaje":"Bienvenidos al chat interno de la empresa.","archivo_nombre":"","archivo_ruta":"","archivo_tipo":""},
            {"fecha":str(datetime.now()),"tipo_chat":"Chat por sucursal","canal":"Sucursal Norte","de_usuario":"norte","de_nombre":"Vendedor Norte","para_usuario":"","mensaje":"Quedan pocas unidades de alimento para gato.","archivo_nombre":"","archivo_ruta":"","archivo_tipo":""},
        ])
        guardar_csv_sistema(chat, CHAT_FILE)


def bloquear_accion():
    st.error("🔒 Esta demo está bloqueada. El cliente puede visualizar el sistema, pero no puede guardar ni modificar datos.")


def guardar_adjunto(uploaded_file, usuario):
    if DEMO_BLOQUEADA:
        return "", "", ""
    if uploaded_file is None:
        return "", "", ""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    nombre_seguro = uploaded_file.name.replace(" ", "_")
    nombre_final = f"{timestamp}_{usuario}_{nombre_seguro}"
    ruta = os.path.join(CHAT_DIR, nombre_final)
    with open(ruta, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return uploaded_file.name, ruta, uploaded_file.type


def mostrar_archivo_chat(nombre, ruta, tipo):
    if not ruta or not os.path.exists(ruta):
        return
    if str(tipo).startswith("image"):
        st.image(ruta, caption=nombre, width=280)
    else:
        with open(ruta, "rb") as f:
            st.download_button(label=f"📎 Descargar {nombre}", data=f, file_name=nombre, mime=tipo if tipo else "application/octet-stream")


# ======================================================
# CSS PREMIUM
# ======================================================
logo_b64 = cargar_logo_base64(LOGO_FILE)

st.markdown(
    """
    <style>
    .stApp { background: #f8fafc; }
    header[data-testid="stHeader"] { background: transparent; }
    section[data-testid="stSidebar"] {
        background: linear-gradient(180deg,#ff7a00 0%,#ff850f 45%,#e66700 100%);
        border-right: 1px solid #ffd3a3;
    }
    section[data-testid="stSidebar"] > div { padding-top: 1rem; }
    .sidebar-logo { text-align: center; padding: 8px 8px 18px 8px; }
    .sidebar-logo img { width: 185px; max-width: 95%; border-radius: 10px; margin-bottom: 12px; border: 3px solid white; box-shadow: 0 8px 22px rgba(0,0,0,0.28); }
    .sidebar-subtitle { color: white; font-size: 12px; font-weight: 900; margin-top: 10px; letter-spacing: 0.4px; }
    .sidebar-divider { margin: 18px auto; width: 82%; height: 1px; background: rgba(255,255,255,0.7); }
    .stRadio > div { gap: 10px; }
    .stRadio label {
        background: rgba(255,255,255,0.14) !important;
        border: 1px solid rgba(255,255,255,0.75) !important;
        border-radius: 8px;
        padding: 10px 14px;
        font-weight: 900;
        color: #ffffff !important;
        box-shadow: 0 4px 12px rgba(0,0,0,0.10);
        width: 100%;
        transition: 0.2s;
    }
    .stRadio label:hover {
        background: #ffffff !important;
        color: #ff6b00 !important;
        border: 1px solid #ffffff !important;
        transform: translateY(-1px);
    }
    .stRadio label p {
        color: inherit !important;
        font-size: 14px !important;
        font-weight: 900 !important;
    }
    section[data-testid="stSidebar"] * {
        text-shadow: 0 1px 2px rgba(0,0,0,0.20);
    }
    section[data-testid="stSidebar"] .stButton button {
        background: #ffffff !important;
        color: #ff6b00 !important;
        text-shadow: none !important;
        border: 1px solid #ffffff !important;
    }
    .main-header { background:white; border:1px solid #edf0f4; border-radius:22px; padding:24px 28px; margin-bottom:20px; box-shadow:0 8px 25px rgba(15,23,42,0.05); }
    .main-title { font-size:34px; font-weight:900; color:#ff6b00; margin-bottom:4px; }
    .main-subtitle { font-size:15px; color:#64748b; font-weight:500; }
    .demo-banner { background:#111827; color:white; border-radius:16px; padding:16px 20px; margin-bottom:18px; font-weight:900; border-left:7px solid #ff7a00; }
    .metric-card { background:white; border:1px solid #edf0f4; border-radius:20px; padding:22px; box-shadow:0 8px 25px rgba(15,23,42,0.06); min-height:145px; }
    .metric-top { display:flex; justify-content:space-between; align-items:center; color:#1f2937; font-size:15px; font-weight:800; }
    .metric-icon { color:#ff7a00; font-size:28px; }
    .metric-value { font-size:28px; font-weight:900; color:#ff6b00; margin-top:24px; }
    .metric-detail { font-size:13px; color:#16a34a; margin-top:8px; font-weight:700; }
    .welcome-card { background:linear-gradient(135deg,#ff7a00 0%,#ff9f1c 100%); color:white; border-radius:24px; padding:34px; box-shadow:0 12px 35px rgba(255,122,0,0.25); margin-bottom:22px; }
    .welcome-title { font-size:38px; font-weight:900; margin-bottom:8px; }
    .welcome-subtitle { font-size:17px; font-weight:600; opacity:.95; }
    .login-logo-box { text-align:center; margin-top:30px; margin-bottom:20px; }
    .login-logo-box img { width:280px; max-width:95%; border-radius:12px; }
    .chat-message { background:white; border:1px solid #edf0f4; border-radius:18px; padding:14px 16px; margin-bottom:12px; box-shadow:0 4px 15px rgba(15,23,42,0.04); }
    .chat-head { font-size:13px; color:#64748b; font-weight:800; margin-bottom:7px; }
    .chat-text { font-size:15px; color:#111827; font-weight:500; }
    .stButton button { background:linear-gradient(135deg,#ff7a00 0%,#ff9f1c 100%); color:white; border:none; border-radius:10px; padding:.7rem 1rem; font-weight:900; box-shadow:0 6px 16px rgba(255,122,0,.28); }
    .stButton button:hover { color:white; background:linear-gradient(135deg,#e66700 0%,#ff7a00 100%); }
    div[data-testid="stDataFrame"] { border-radius:16px; overflow:hidden; }
    .footer-sidebar { font-size:11px; color:#ffffff; margin-top:35px; text-align:center; font-weight:900; }
    </style>
    """,
    unsafe_allow_html=True
)

# ======================================================
# INICIO DEL SISTEMA
# ======================================================
crear_archivos_iniciales()
stock_df = cargar_csv(STOCK_FILE)
ventas_df = cargar_csv(VENTAS_FILE)
compras_df = cargar_csv(COMPRAS_FILE)
clientes_df = cargar_csv(CLIENTES_FILE)
perdidas_df = cargar_csv(PERDIDAS_FILE)
chat_df = cargar_csv(CHAT_FILE)

if "logueado" not in st.session_state:
    st.session_state.logueado = False
if "usuario" not in st.session_state:
    st.session_state.usuario = None

# ======================================================
# LOGIN
# ======================================================
if not st.session_state.logueado:
    col1, col2, col3 = st.columns([1, 1.05, 1])
    with col2:
        if logo_b64:
            st.markdown(f"""<div class="login-logo-box"><img src="data:image/png;base64,{logo_b64}"></div>""", unsafe_allow_html=True)
        else:
            st.markdown("<h1 style='text-align:center;color:#ff6b00;'>🐾 PetShop Manager Pro</h1>", unsafe_allow_html=True)

        st.title("🔐 Iniciar sesión")
        st.caption("Modo demo bloqueado para clientes.")
        usuario = st.text_input("Usuario", value="demo")
        password = st.text_input("Contraseña", type="password", value="1234")
        st.info("Demo bloqueada: demo / 1234")

        if st.button("Ingresar al sistema"):
            if usuario in USUARIOS and USUARIOS[usuario]["password"] == password:
                st.session_state.logueado = True
                st.session_state.usuario = usuario
                st.rerun()
            else:
                st.error("Usuario o contraseña incorrectos.")
    st.stop()

usuario_actual = st.session_state.usuario
rol_actual = USUARIOS[usuario_actual]["rol"]
sucursal_usuario = USUARIOS[usuario_actual]["sucursal"]
nombre_actual = USUARIOS[usuario_actual]["nombre"]

# ======================================================
# SIDEBAR
# ======================================================
with st.sidebar:
    if logo_b64:
        st.markdown(f"""
        <div class="sidebar-logo">
            <img src="data:image/png;base64,{logo_b64}">
            <div class="sidebar-subtitle">Sistema de gestión integral</div>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.markdown("""
        <div class="sidebar-logo">
            <div style="font-size:22px;font-weight:900;color:white;">🐾 PetShop Manager Pro</div>
            <div class="sidebar-subtitle">Sistema de gestión integral</div>
        </div>
        """, unsafe_allow_html=True)

    st.markdown('<div class="sidebar-divider"></div>', unsafe_allow_html=True)

    opciones = [
        "Inicio", "Ventas", "Stock", "Sucursales", "Clientes", "Compras",
        "Ganancias y pérdidas", "Reportes", "Chat interno", "Importar Excel",
        "Exportar datos", "Configuración"
    ]
    menu = st.radio("Menú", opciones, label_visibility="collapsed")

    st.markdown('<div class="sidebar-divider"></div>', unsafe_allow_html=True)
    st.markdown(f"""
    <div style="background:rgba(255,255,255,0.14);border:1px solid rgba(255,255,255,0.70);border-radius:10px;padding:13px;color:white;font-weight:900;">
        🧾 Usuario: {usuario_actual}<br>
        🔐 Rol: {rol_actual}<br>
        🏪 Sucursal: {sucursal_usuario}
    </div>
    """, unsafe_allow_html=True)
    st.markdown("""
    <div style="background:rgba(255,255,255,0.14);border:1px solid rgba(255,255,255,0.70);border-radius:10px;padding:13px;color:white;font-weight:900;margin-top:12px;">
        📅 Ciclo comercial<br>
        Mayo 2026
    </div>
    """, unsafe_allow_html=True)
    if st.button("Cerrar sesión"):
        st.session_state.logueado = False
        st.session_state.usuario = None
        st.rerun()
    st.markdown("""
    <div class="footer-sidebar">
        🐾 PetShop Manager Pro<br>
        © 2026 La Casa del Mascotero
    </div>
    """, unsafe_allow_html=True)

st.markdown(f"""
<div class="main-header">
    <div class="main-title">🐾 PetShop Manager Pro</div>
    <div class="main-subtitle">Bienvenido. Sistema de gestión integral para pet shops con múltiples sucursales.</div>
</div>
""", unsafe_allow_html=True)

if DEMO_BLOQUEADA:
    st.markdown("""
    <div class="demo-banner">
        🔒 MODO DEMO BLOQUEADO — El cliente puede navegar y visualizar el sistema, pero no puede guardar, modificar, importar ni enviar mensajes.
    </div>
    """, unsafe_allow_html=True)


def filtrar_por_rol(df):
    if df.empty:
        return df
    if rol_actual in ["Administrador", "Demo"]:
        return df
    if "sucursal" in df.columns:
        return df[df["sucursal"] == sucursal_usuario]
    return df

ventas_visibles = filtrar_por_rol(ventas_df)
stock_visible = filtrar_por_rol(stock_df)
compras_visibles = filtrar_por_rol(compras_df)
perdidas_visibles = filtrar_por_rol(perdidas_df)

# ======================================================
# PÁGINAS
# ======================================================
if menu == "Inicio":
    st.markdown("""
    <div class="welcome-card">
        <div class="welcome-title">Panel de control empresarial</div>
        <div class="welcome-subtitle">Controlá ventas, stock, rentabilidad, pérdidas, comunicación interna y rendimiento de sucursales.</div>
    </div>
    """, unsafe_allow_html=True)

    if not ventas_visibles.empty:
        vtmp = ventas_visibles.copy()
        vtmp["fecha"] = pd.to_datetime(vtmp["fecha"], errors="coerce")
        hoy = date.today()
        fecha_hoy = pd.to_datetime(hoy)
        ventas_hoy = vtmp[vtmp["fecha"].dt.date == hoy]
        ventas_mes = vtmp[(vtmp["fecha"].dt.month == fecha_hoy.month) & (vtmp["fecha"].dt.year == fecha_hoy.year)]
        ventas_anio = vtmp[vtmp["fecha"].dt.year == fecha_hoy.year]
        fact_hoy = ventas_hoy["total"].sum()
        fact_mes = ventas_mes["total"].sum()
        fact_anio = ventas_anio["total"].sum()
        ticket = vtmp["total"].mean()
        ganancia_mes = ventas_mes["ganancia"].sum()
    else:
        fact_hoy = fact_mes = fact_anio = ticket = ganancia_mes = 0

    perdida_total = perdidas_visibles["perdida_total"].sum() if not perdidas_visibles.empty else 0
    stock_valorizado = (stock_visible["stock"] * stock_visible["precio_compra"]).sum() if not stock_visible.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1: metric_card("Ventas de hoy", pesos(fact_hoy), "🛒", "Demo visual")
    with c2: metric_card("Ventas del mes", pesos(fact_mes), "📈", "Ciclo comercial")
    with c3: metric_card("Ganancia mensual", pesos(ganancia_mes), "💰", "Estimación bruta")
    with c4: metric_card("Ticket promedio", pesos(ticket), "🎟️", "Promedio por venta")
    c5, c6, c7, c8 = st.columns(4)
    with c5: metric_card("Ventas del año", pesos(fact_anio), "📅", "Acumulado anual")
    with c6: metric_card("Stock valorizado", pesos(stock_valorizado), "📦", "Costo inventario")
    with c7: metric_card("Pérdidas registradas", pesos(perdida_total), "⚠️", "Control operativo")
    with c8: metric_card("Mensajes internos", len(chat_df), "💬", "Comunicación")

    st.divider()
    col1, col2 = st.columns([2, 1])
    with col1:
        st.subheader("📈 Evolución de ventas")
        if not ventas_visibles.empty:
            chart_df = ventas_visibles.copy()
            chart_df["fecha"] = pd.to_datetime(chart_df["fecha"], errors="coerce")
            chart_df["dia"] = chart_df["fecha"].dt.date
            ventas_dia = chart_df.groupby("dia")["total"].sum().reset_index()
            st.line_chart(ventas_dia.set_index("dia"))
        else:
            st.info("Todavía no hay ventas cargadas.")
    with col2:
        st.subheader("🏪 Ranking de sucursales")
        if not ventas_df.empty:
            ranking = ventas_df.groupby("sucursal")["total"].sum().reset_index().sort_values("total", ascending=False)
            st.dataframe(ranking, use_container_width=True, hide_index=True)
        else:
            st.info("Sin datos de ventas.")

    col3, col4, col5 = st.columns(3)
    with col3:
        st.subheader("🔥 Productos más vendidos")
        if not ventas_visibles.empty:
            top = ventas_visibles.groupby("producto")["cantidad"].sum().reset_index().sort_values("cantidad", ascending=False).head(8)
            st.dataframe(top, use_container_width=True, hide_index=True)
        else:
            st.info("Sin ventas.")
    with col4:
        st.subheader("⚠️ Bajo stock")
        bajo = stock_visible[stock_visible["stock"] <= stock_visible["stock_minimo"]] if not stock_visible.empty else pd.DataFrame()
        if not bajo.empty:
            st.dataframe(bajo[["producto", "sucursal", "stock", "stock_minimo"]], use_container_width=True, hide_index=True)
        else:
            st.success("Stock en buen estado.")
    with col5:
        st.subheader("💳 Medios de pago")
        if not ventas_visibles.empty:
            medios = ventas_visibles.groupby("medio_pago")["total"].sum().reset_index()
            st.dataframe(medios, use_container_width=True, hide_index=True)
        else:
            st.info("Sin pagos registrados.")

elif menu == "Ventas":
    st.header("🛒 Gestión de ventas")
    st.warning("🔒 Demo bloqueada: el formulario se puede visualizar, pero no guarda ventas.")

    if stock_visible.empty:
        st.warning("Primero cargá productos en stock.")
    else:
        with st.form("form_venta"):
            c1, c2, c3 = st.columns(3)
            sucursal = c1.selectbox("Sucursal", SUCURSALES)
            stock_sucursal = stock_df[stock_df["sucursal"] == sucursal]
            producto_nombre = c2.selectbox("Producto", stock_sucursal["producto"].unique()) if not stock_sucursal.empty else None
            cantidad = c3.number_input("Cantidad", min_value=1, step=1)
            c4, c5, c6 = st.columns(3)
            if producto_nombre:
                row = stock_sucursal[stock_sucursal["producto"] == producto_nombre].iloc[0]
                precio_unitario = c4.number_input("Precio unitario", value=float(row["precio_venta"]), min_value=0.0, step=100.0)
                precio_compra = float(row["precio_compra"])
                stock_actual = int(row["stock"])
            else:
                precio_unitario = c4.number_input("Precio unitario", min_value=0.0, step=100.0)
                precio_compra = 0
                stock_actual = 0
            descuento = c5.number_input("Descuento", min_value=0.0, value=0.0, step=100.0)
            medio_pago = c6.selectbox("Medio de pago", MEDIOS_PAGO)
            c7, c8, c9 = st.columns(3)
            c7.text_input("Vendedor", value=usuario_actual)
            c8.text_input("Cliente")
            c9.text_input("Comprobante")
            total = cantidad * precio_unitario - descuento
            ganancia = total - (cantidad * precio_compra)
            margen = (ganancia / total * 100) if total > 0 else 0
            st.info(f"Total: {pesos(total)} | Ganancia: {pesos(ganancia)} | Margen: {margen:.2f}% | Stock actual: {stock_actual}")
            if st.form_submit_button("Guardar venta"):
                bloquear_accion()

    st.subheader("Historial de ventas")
    st.dataframe(ventas_visibles, use_container_width=True, hide_index=True)

elif menu == "Stock":
    st.header("📦 Gestión de stock")
    c1, c2, c3 = st.columns(3)
    filtro_sucursal = c1.selectbox("Sucursal", ["Todas"] + SUCURSALES)
    filtro_categoria = c2.selectbox("Categoría", ["Todas"] + CATEGORIAS)
    buscar = c3.text_input("Buscar producto")
    df = stock_visible.copy()
    if filtro_sucursal != "Todas":
        df = df[df["sucursal"] == filtro_sucursal]
    if filtro_categoria != "Todas":
        df = df[df["categoria"] == filtro_categoria]
    if buscar:
        df = df[df["producto"].str.contains(buscar, case=False, na=False)]
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("➕ Agregar producto")
    st.warning("🔒 Demo bloqueada: no se pueden guardar productos.")
    with st.form("form_stock"):
        c1, c2, c3 = st.columns(3)
        c1.text_input("Código")
        c2.text_input("Producto")
        c3.selectbox("Categoría", CATEGORIAS)
        c4, c5, c6 = st.columns(3)
        c4.text_input("Marca")
        c5.text_input("Proveedor")
        c6.selectbox("Sucursal", SUCURSALES)
        c7, c8, c9 = st.columns(3)
        c7.number_input("Precio compra", min_value=0.0, step=100.0)
        c8.number_input("Precio venta", min_value=0.0, step=100.0)
        c9.number_input("Stock", min_value=0, step=1)
        if st.form_submit_button("Guardar producto"):
            bloquear_accion()

elif menu == "Sucursales":
    st.header("🏪 Sucursales")
    sucursal = st.selectbox("Seleccionar sucursal", SUCURSALES)
    v = ventas_df[ventas_df["sucursal"] == sucursal] if not ventas_df.empty else pd.DataFrame()
    s = stock_df[stock_df["sucursal"] == sucursal]
    p = perdidas_df[perdidas_df["sucursal"] == sucursal] if not perdidas_df.empty else pd.DataFrame()
    fact = v["total"].sum() if not v.empty else 0
    gan = v["ganancia"].sum() if not v.empty else 0
    per = p["perdida_total"].sum() if not p.empty else 0
    inv = (s["stock"] * s["precio_compra"]).sum() if not s.empty else 0
    c1, c2, c3, c4 = st.columns(4)
    with c1: metric_card("Facturación", pesos(fact), "💰")
    with c2: metric_card("Ganancia", pesos(gan), "📈")
    with c3: metric_card("Pérdidas", pesos(per), "⚠️")
    with c4: metric_card("Inventario", pesos(inv), "📦")
    st.subheader("Ventas")
    st.dataframe(v, use_container_width=True, hide_index=True)
    st.subheader("Stock")
    st.dataframe(s, use_container_width=True, hide_index=True)

elif menu == "Clientes":
    st.header("👥 Clientes")
    st.warning("🔒 Demo bloqueada: el formulario se puede visualizar, pero no guarda clientes.")
    with st.form("form_cliente"):
        c1, c2, c3 = st.columns(3)
        c1.text_input("Nombre")
        c2.text_input("Teléfono")
        c3.text_input("Email")
        c4, c5, c6 = st.columns(3)
        c4.text_input("Mascota")
        c5.selectbox("Tipo de mascota", ["Perro", "Gato", "Ave", "Roedor", "Otro"])
        c6.text_input("Observaciones")
        if st.form_submit_button("Guardar cliente"):
            bloquear_accion()
    st.dataframe(clientes_df, use_container_width=True, hide_index=True)

elif menu == "Compras":
    st.header("🚚 Compras a proveedores")
    st.warning("🔒 Demo bloqueada: no se pueden guardar compras ni actualizar stock.")
    with st.form("form_compra"):
        c1, c2, c3 = st.columns(3)
        c1.text_input("Proveedor")
        c2.text_input("Código")
        c3.text_input("Producto")
        c4, c5, c6 = st.columns(3)
        c4.selectbox("Categoría", CATEGORIAS)
        c5.selectbox("Sucursal destino", SUCURSALES)
        c6.number_input("Cantidad", min_value=1, step=1)
        c7, c8 = st.columns(2)
        c7.number_input("Precio compra unitario", min_value=0.0, step=100.0)
        c8.number_input("Precio venta sugerido", min_value=0.0, step=100.0)
        if st.form_submit_button("Guardar compra"):
            bloquear_accion()
    st.dataframe(compras_visibles, use_container_width=True, hide_index=True)

elif menu == "Ganancias y pérdidas":
    st.header("💰 Ganancias y pérdidas")
    total_fact = ventas_df["total"].sum() if not ventas_df.empty else 0
    total_costo = ventas_df["costo_total"].sum() if not ventas_df.empty else 0
    total_gan = ventas_df["ganancia"].sum() if not ventas_df.empty else 0
    total_perd = perdidas_df["perdida_total"].sum() if not perdidas_df.empty else 0
    c1, c2, c3, c4 = st.columns(4)
    with c1: metric_card("Facturación", pesos(total_fact), "💰")
    with c2: metric_card("Costo vendido", pesos(total_costo), "📦")
    with c3: metric_card("Ganancia bruta", pesos(total_gan), "📈")
    with c4: metric_card("Resultado neto", pesos(total_gan - total_perd), "🏆")
    st.subheader("Historial de pérdidas")
    st.dataframe(perdidas_df, use_container_width=True, hide_index=True)

elif menu == "Reportes":
    st.header("📊 Reportes avanzados")
    if ventas_visibles.empty:
        st.info("No hay ventas para analizar.")
    else:
        reporte = ventas_visibles.copy()
        reporte["fecha"] = pd.to_datetime(reporte["fecha"], errors="coerce")
        reporte["dia"] = reporte["fecha"].dt.date
        reporte["mes"] = reporte["fecha"].dt.to_period("M").astype(str)
        st.subheader("Ventas por día")
        ventas_dia = reporte.groupby("dia")["total"].sum().reset_index()
        st.line_chart(ventas_dia.set_index("dia"))
        st.subheader("Ventas por mes")
        ventas_mes = reporte.groupby("mes")["total"].sum().reset_index()
        st.bar_chart(ventas_mes.set_index("mes"))
        st.subheader("Productos más rentables")
        rentables = reporte.groupby("producto")["ganancia"].sum().reset_index().sort_values("ganancia", ascending=False)
        st.dataframe(rentables, use_container_width=True, hide_index=True)
        st.subheader("Ventas por categoría")
        cat = reporte.groupby("categoria")["total"].sum().reset_index()
        st.bar_chart(cat.set_index("categoria"))
        st.subheader("Medios de pago")
        medios = reporte.groupby("medio_pago")["total"].sum().reset_index()
        st.dataframe(medios, use_container_width=True, hide_index=True)

elif menu == "Chat interno":
    st.header("💬 Chat interno")
    st.warning("🔒 Demo bloqueada: el chat se puede visualizar, pero no permite enviar mensajes ni adjuntos.")
    chat_df = cargar_csv(CHAT_FILE)
    tipo_chat = st.radio("Seleccionar tipo de chat", ["Chat general", "Chat por sucursal", "Chat privado"], horizontal=True)
    canal = "general"
    para_usuario = ""
    if tipo_chat == "Chat general":
        st.subheader("🌐 Chat general de la empresa")
        canal = "general"
    elif tipo_chat == "Chat por sucursal":
        canal = st.selectbox("Sucursal", SUCURSALES)
        st.subheader(f"🏪 Chat de {canal}")
    else:
        usuarios_disponibles = [u for u in USUARIOS.keys() if u != usuario_actual]
        para_usuario = st.selectbox("Hablar con", usuarios_disponibles)
        participantes = sorted([usuario_actual, para_usuario])
        canal = f"privado_{participantes[0]}_{participantes[1]}"
        st.subheader(f"🔒 Chat privado con {para_usuario}")
    st.divider()
    with st.form("form_chat", clear_on_submit=True):
        st.text_area("Escribir mensaje", height=90)
        st.file_uploader("Adjuntar archivo", type=["png", "jpg", "jpeg", "pdf", "xlsx", "xls", "docx", "txt"])
        if st.form_submit_button("Enviar mensaje"):
            bloquear_accion()
    st.divider()
    st.subheader("📨 Mensajes")
    if chat_df.empty:
        st.info("Todavía no hay mensajes.")
    else:
        mensajes = chat_df[chat_df["canal"] == canal].copy()
        if not mensajes.empty:
            mensajes["fecha"] = pd.to_datetime(mensajes["fecha"], errors="coerce")
            mensajes = mensajes.sort_values("fecha", ascending=False)
            for _, msg in mensajes.iterrows():
                fecha_txt = msg["fecha"].strftime("%d/%m/%Y %H:%M") if pd.notnull(msg["fecha"]) else ""
                texto = msg["mensaje"] if pd.notnull(msg["mensaje"]) else ""
                st.markdown(f"""
                <div class="chat-message">
                    <div class="chat-head">👤 {msg['de_nombre']} · {msg['de_usuario']} · {fecha_txt}</div>
                    <div class="chat-text">{texto}</div>
                </div>
                """, unsafe_allow_html=True)
                if pd.notnull(msg.get("archivo_ruta", "")) and str(msg.get("archivo_ruta", "")) != "":
                    mostrar_archivo_chat(msg.get("archivo_nombre", ""), msg.get("archivo_ruta", ""), msg.get("archivo_tipo", ""))
        else:
            st.info("No hay mensajes en este chat todavía.")

elif menu == "Importar Excel":
    st.header("📥 Importar Excel")
    st.error("🔒 Importación bloqueada en modo demo.")
    st.file_uploader("Subir Excel", type=["xlsx", "xls"], disabled=True)

elif menu == "Exportar datos":
    st.header("📤 Exportar datos")
    st.info("En esta demo se permite descargar un reporte de muestra para presentación comercial.")
    excel_data = exportar_excel({"Ventas": ventas_df, "Stock": stock_df, "Compras": compras_df, "Clientes": clientes_df, "Perdidas": perdidas_df, "Chat": chat_df})
    st.download_button("Descargar Excel demo", data=excel_data, file_name="petshop_manager_demo.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

elif menu == "Configuración":
    st.header("⚙️ Configuración del sistema")
    st.error("🔒 Configuración bloqueada para clientes en modo demo.")
    usuarios_demo = pd.DataFrame([
        {"Usuario":"demo","Contraseña":"1234","Rol":"Demo","Acceso":"Visualización bloqueada"},
        {"Usuario":"admin","Contraseña":"1234","Rol":"Administrador","Acceso":"Solo para presentación"},
        {"Usuario":"encargado","Contraseña":"1234","Rol":"Encargado","Acceso":"Solo para presentación"},
        {"Usuario":"vendedor","Contraseña":"1234","Rol":"Vendedor","Acceso":"Solo para presentación"},
        {"Usuario":"norte","Contraseña":"1234","Rol":"Vendedor","Acceso":"Solo para presentación"},
        {"Usuario":"sur","Contraseña":"1234","Rol":"Vendedor","Acceso":"Solo para presentación"},
    ])
    st.dataframe(usuarios_demo, use_container_width=True, hide_index=True)
