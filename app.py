import streamlit as st
import pandas as pd
import os
from datetime import datetime, date
from io import BytesIO
import base64

st.set_page_config(
    page_title="PetShop Manager Pro",
    page_icon="🐾",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================
# ARCHIVOS / CONFIG
# =========================

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
    "Alimento para perros",
    "Alimento para gatos",
    "Snacks",
    "Piedras sanitarias",
    "Arena para gatos",
    "Juguetes",
    "Correas",
    "Collares",
    "Camas",
    "Comederos",
    "Bebederos",
    "Shampoo",
    "Antipulgas",
    "Medicamentos",
    "Accesorios",
    "Ropa para mascotas",
    "Productos de higiene",
    "Otros"
]

MEDIOS_PAGO = ["Efectivo", "Débito", "Crédito", "Transferencia", "Mercado Pago"]

USUARIOS = {
    "admin": {
        "password": "1234",
        "rol": "Administrador",
        "sucursal": "Todas",
        "nombre": "Dueño"
    },
    "encargado": {
        "password": "1234",
        "rol": "Encargado",
        "sucursal": "Casa Central",
        "nombre": "Encargado Casa Central"
    },
    "vendedor": {
        "password": "1234",
        "rol": "Vendedor",
        "sucursal": "Casa Central",
        "nombre": "Vendedor Casa Central"
    },
    "norte": {
        "password": "1234",
        "rol": "Vendedor",
        "sucursal": "Sucursal Norte",
        "nombre": "Vendedor Norte"
    },
    "sur": {
        "password": "1234",
        "rol": "Vendedor",
        "sucursal": "Sucursal Sur",
        "nombre": "Vendedor Sur"
    },
}


# =========================
# FUNCIONES
# =========================

def cargar_logo_base64(path):
    if os.path.exists(path):
        with open(path, "rb") as img:
            return base64.b64encode(img.read()).decode()
    return None


def crear_archivos_iniciales():
    if not os.path.exists(STOCK_FILE):
        stock = pd.DataFrame([
            {
                "codigo": "DOG001",
                "producto": "Royal Canin Mini Adult 3kg",
                "categoria": "Alimento para perros",
                "marca": "Royal Canin",
                "proveedor": "Proveedor A",
                "sucursal": "Casa Central",
                "precio_compra": 15000,
                "precio_venta": 23000,
                "stock": 12,
                "stock_minimo": 5,
                "fecha_ingreso": str(date.today()),
                "fecha_vencimiento": ""
            },
            {
                "codigo": "CAT001",
                "producto": "Piedras Sanitarias 4kg",
                "categoria": "Piedras sanitarias",
                "marca": "MichiClean",
                "proveedor": "Proveedor B",
                "sucursal": "Sucursal Norte",
                "precio_compra": 2500,
                "precio_venta": 4300,
                "stock": 25,
                "stock_minimo": 8,
                "fecha_ingreso": str(date.today()),
                "fecha_vencimiento": ""
            },
            {
                "codigo": "ACC001",
                "producto": "Correa Retráctil Talla M",
                "categoria": "Correas",
                "marca": "PetFlex",
                "proveedor": "Proveedor C",
                "sucursal": "Sucursal Sur",
                "precio_compra": 6000,
                "precio_venta": 11500,
                "stock": 4,
                "stock_minimo": 5,
                "fecha_ingreso": str(date.today()),
                "fecha_vencimiento": ""
            }
        ])
        stock.to_csv(STOCK_FILE, index=False)

    if not os.path.exists(VENTAS_FILE):
        pd.DataFrame(columns=[
            "fecha", "sucursal", "codigo", "producto", "categoria", "cantidad",
            "precio_unitario", "descuento", "total", "costo_total", "ganancia",
            "medio_pago", "vendedor", "cliente", "comprobante"
        ]).to_csv(VENTAS_FILE, index=False)

    if not os.path.exists(COMPRAS_FILE):
        pd.DataFrame(columns=[
            "fecha", "proveedor", "codigo", "producto", "categoria", "sucursal",
            "cantidad", "precio_compra_unitario", "costo_total"
        ]).to_csv(COMPRAS_FILE, index=False)

    if not os.path.exists(CLIENTES_FILE):
        pd.DataFrame(columns=[
            "nombre", "telefono", "mascota", "tipo_mascota", "email", "observaciones"
        ]).to_csv(CLIENTES_FILE, index=False)

    if not os.path.exists(PERDIDAS_FILE):
        pd.DataFrame(columns=[
            "fecha", "sucursal", "codigo", "producto", "motivo", "cantidad",
            "costo_unitario", "perdida_total", "observaciones"
        ]).to_csv(PERDIDAS_FILE, index=False)

    if not os.path.exists(CHAT_FILE):
        pd.DataFrame(columns=[
            "fecha", "tipo_chat", "canal", "de_usuario", "de_nombre",
            "para_usuario", "mensaje", "archivo_nombre", "archivo_ruta", "archivo_tipo"
        ]).to_csv(CHAT_FILE, index=False)


def cargar_csv(path):
    if os.path.exists(path):
        return pd.read_csv(path)
    return pd.DataFrame()


def guardar_csv(df, path):
    df.to_csv(path, index=False)


def pesos(valor):
    try:
        return f"$ {float(valor):,.0f}".replace(",", ".")
    except:
        return "$ 0"


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


def guardar_adjunto(uploaded_file, usuario):
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
            st.download_button(
                label=f"📎 Descargar {nombre}",
                data=f,
                file_name=nombre,
                mime=tipo if tipo else "application/octet-stream"
            )


# =========================
# ESTILOS
# =========================

logo_b64 = cargar_logo_base64(LOGO_FILE)

st.markdown(
    """
    <style>
    .stApp {
        background: #f8fafc;
    }

    header[data-testid="stHeader"] {
        background: transparent;
    }

    section[data-testid="stSidebar"] {
        background: linear-gradient(
            180deg,
            #ff7a00 0%,
            #ff8f1f 38%,
            #ff9f1c 70%,
            #ffffff 100%
        );
        border-right: 1px solid #ffd3a3;
    }

    section[data-testid="stSidebar"] > div {
        padding-top: 1rem;
    }

    .sidebar-logo {
        text-align: center;
        padding: 8px 8px 18px 8px;
    }

    .sidebar-logo img {
        width: 185px;
        max-width: 95%;
        border-radius: 10px;
        margin-bottom: 12px;
        border: 3px solid white;
        box-shadow: 0 8px 22px rgba(0,0,0,0.28);
    }

    .sidebar-subtitle {
        color: white;
        font-size: 12px;
        font-weight: 900;
        margin-top: 10px;
        letter-spacing: 0.4px;
    }

    .sidebar-divider {
        margin: 18px auto;
        width: 82%;
        height: 1px;
        background: rgba(255,255,255,0.7);
    }

    .stRadio > div {
        gap: 10px;
    }

    .stRadio label {
        background: linear-gradient(135deg, #ffffff 0%, #fff5eb 100%);
        border: 1px solid #ffd2a8;
        border-radius: 8px;
        padding: 10px 14px;
        font-weight: 800;
        color: #1f2937;
        box-shadow: 0 4px 12px rgba(0,0,0,0.08);
        width: 100%;
        transition: 0.2s;
    }

    .stRadio label:hover {
        background: linear-gradient(135deg, #ff7a00 0%, #ff9f1c 100%);
        color: white;
        border: 1px solid #ff7a00;
        transform: translateY(-1px);
    }

    .stRadio label p {
        font-size: 14px !important;
        font-weight: 800 !important;
    }

    .main-header {
        background: white;
        border: 1px solid #edf0f4;
        border-radius: 22px;
        padding: 24px 28px;
        margin-bottom: 20px;
        box-shadow: 0 8px 25px rgba(15,23,42,0.05);
    }

    .main-title {
        font-size: 34px;
        font-weight: 900;
        color: #ff6b00;
        margin-bottom: 4px;
    }

    .main-subtitle {
        font-size: 15px;
        color: #64748b;
        font-weight: 500;
    }

    .metric-card {
        background: white;
        border: 1px solid #edf0f4;
        border-radius: 20px;
        padding: 22px;
        box-shadow: 0 8px 25px rgba(15,23,42,0.06);
        min-height: 145px;
    }

    .metric-top {
        display: flex;
        justify-content: space-between;
        align-items: center;
        color: #1f2937;
        font-size: 15px;
        font-weight: 800;
    }

    .metric-icon {
        color: #ff7a00;
        font-size: 28px;
    }

    .metric-value {
        font-size: 28px;
        font-weight: 900;
        color: #ff6b00;
        margin-top: 24px;
    }

    .metric-detail {
        font-size: 13px;
        color: #16a34a;
        margin-top: 8px;
        font-weight: 700;
    }

    .welcome-card {
        background: linear-gradient(135deg, #ff7a00 0%, #ff9f1c 100%);
        color: white;
        border-radius: 24px;
        padding: 34px;
        box-shadow: 0 12px 35px rgba(255,122,0,0.25);
        margin-bottom: 22px;
    }

    .welcome-title {
        font-size: 38px;
        font-weight: 900;
        margin-bottom: 8px;
    }

    .welcome-subtitle {
        font-size: 17px;
        font-weight: 600;
        opacity: 0.95;
    }

    .login-logo-box {
        text-align: center;
        margin-top: 30px;
        margin-bottom: 20px;
    }

    .login-logo-box img {
        width: 260px;
        max-width: 95%;
        border-radius: 12px;
    }

    .chat-message {
        background: white;
        border: 1px solid #edf0f4;
        border-radius: 18px;
        padding: 14px 16px;
        margin-bottom: 12px;
        box-shadow: 0 4px 15px rgba(15,23,42,0.04);
    }

    .chat-head {
        font-size: 13px;
        color: #64748b;
        font-weight: 800;
        margin-bottom: 7px;
    }

    .chat-text {
        font-size: 15px;
        color: #111827;
        font-weight: 500;
    }

    .stButton button {
        background: linear-gradient(135deg, #ff7a00 0%, #ff9f1c 100%);
        color: white;
        border: none;
        border-radius: 10px;
        padding: 0.7rem 1rem;
        font-weight: 900;
        box-shadow: 0 6px 16px rgba(255,122,0,0.28);
    }

    .stButton button:hover {
        color: white;
        background: linear-gradient(135deg, #e66700 0%, #ff7a00 100%);
    }

    div[data-testid="stDataFrame"] {
        border-radius: 16px;
        overflow: hidden;
    }

    .footer-sidebar {
        font-size: 11px;
        color: #3f3f46;
        margin-top: 35px;
        text-align: center;
        font-weight: 800;
    }
    </style>
    """,
    unsafe_allow_html=True
)

# =========================
# CARGA DE DATOS
# =========================

crear_archivos_iniciales()

stock_df = cargar_csv(STOCK_FILE)
ventas_df = cargar_csv(VENTAS_FILE)
compras_df = cargar_csv(COMPRAS_FILE)
clientes_df = cargar_csv(CLIENTES_FILE)
perdidas_df = cargar_csv(PERDIDAS_FILE)
chat_df = cargar_csv(CHAT_FILE)

# =========================
# LOGIN
# =========================

if "logueado" not in st.session_state:
    st.session_state.logueado = False

if "usuario" not in st.session_state:
    st.session_state.usuario = None

if not st.session_state.logueado:
    col1, col2, col3 = st.columns([1, 1.05, 1])

    with col2:
        if logo_b64:
            st.markdown(
                f"""
                <div class="login-logo-box">
                    <img src="data:image/png;base64,{logo_b64}">
                </div>
                """,
                unsafe_allow_html=True
            )

        st.title("🔐 Iniciar sesión")
        st.caption("Ingresá al sistema de gestión del pet shop.")

        usuario = st.text_input("Usuario")
        password = st.text_input("Contraseña", type="password")

        st.info("Demo: admin / 1234")

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

# =========================
# SIDEBAR
# =========================

with st.sidebar:
    if logo_b64:
        st.markdown(
            f"""
            <div class="sidebar-logo">
                <img src="data:image/png;base64,{logo_b64}">
                <div class="sidebar-subtitle">Sistema de gestión integral</div>
            </div>
            """,
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            """
            <div class="sidebar-logo">
                <div style="font-size:22px;font-weight:900;color:white;">🐾 PetShop Manager Pro</div>
                <div class="sidebar-subtitle">Sistema de gestión integral</div>
            </div>
            """,
            unsafe_allow_html=True
        )

    st.markdown('<div class="sidebar-divider"></div>', unsafe_allow_html=True)

    if rol_actual == "Vendedor":
        opciones = ["Inicio", "Ventas", "Clientes", "Chat interno"]
    elif rol_actual == "Encargado":
        opciones = ["Inicio", "Ventas", "Stock", "Clientes", "Compras", "Reportes", "Chat interno"]
    else:
        opciones = [
            "Inicio",
            "Ventas",
            "Stock",
            "Sucursales",
            "Clientes",
            "Compras",
            "Ganancias y pérdidas",
            "Reportes",
            "Chat interno",
            "Importar Excel",
            "Exportar datos",
            "Configuración"
        ]

    menu = st.radio("Menú", opciones, label_visibility="collapsed")

    st.markdown('<div class="sidebar-divider"></div>', unsafe_allow_html=True)

    st.markdown(
        f"""
        <div style="background:white;border-radius:10px;padding:13px;color:#1f2937;font-weight:800;">
            🧾 Usuario: {usuario_actual}<br>
            🔐 Rol: {rol_actual}<br>
            🏪 Sucursal: {sucursal_usuario}
        </div>
        """,
        unsafe_allow_html=True
    )

    st.markdown(
        """
        <div style="background:white;border-radius:10px;padding:13px;color:#1f2937;font-weight:800;margin-top:12px;">
            📅 Ciclo comercial<br>
            Mayo 2026
        </div>
        """,
        unsafe_allow_html=True
    )

    if st.button("Cerrar sesión"):
        st.session_state.logueado = False
        st.session_state.usuario = None
        st.rerun()

    st.markdown(
        """
        <div class="footer-sidebar">
            🐾 PetShop Manager Pro<br>
            © 2026 La Casa del Mascotero
        </div>
        """,
        unsafe_allow_html=True
    )

# =========================
# HEADER
# =========================

st.markdown(
    f"""
    <div class="main-header">
        <div class="main-title">🐾 PetShop Manager Pro</div>
        <div class="main-subtitle">Bienvenido, {rol_actual}. Sistema de gestión integral para pet shops con múltiples sucursales.</div>
    </div>
    """,
    unsafe_allow_html=True
)

# =========================
# FILTROS POR ROL
# =========================

def filtrar_por_rol(df):
    if df.empty:
        return df
    if rol_actual == "Administrador":
        return df
    if "sucursal" in df.columns:
        return df[df["sucursal"] == sucursal_usuario]
    return df


ventas_visibles = filtrar_por_rol(ventas_df)
stock_visible = filtrar_por_rol(stock_df)
compras_visibles = filtrar_por_rol(compras_df)
perdidas_visibles = filtrar_por_rol(perdidas_df)

# =========================
# INICIO
# =========================

if menu == "Inicio":
    st.markdown(
        """
        <div class="welcome-card">
            <div class="welcome-title">Panel de control empresarial</div>
            <div class="welcome-subtitle">
                Controlá ventas, stock, rentabilidad, pérdidas, comunicación interna y rendimiento de sucursales.
            </div>
        </div>
        """,
        unsafe_allow_html=True
    )

    if not ventas_visibles.empty:
        ventas_visibles["fecha"] = pd.to_datetime(ventas_visibles["fecha"], errors="coerce")
        hoy = date.today()
        fecha_hoy = pd.to_datetime(hoy)

        ventas_hoy = ventas_visibles[ventas_visibles["fecha"].dt.date == hoy]
        ventas_mes = ventas_visibles[
            (ventas_visibles["fecha"].dt.month == fecha_hoy.month) &
            (ventas_visibles["fecha"].dt.year == fecha_hoy.year)
        ]
        ventas_anio = ventas_visibles[ventas_visibles["fecha"].dt.year == fecha_hoy.year]

        fact_hoy = ventas_hoy["total"].sum()
        fact_mes = ventas_mes["total"].sum()
        fact_anio = ventas_anio["total"].sum()
        ticket = ventas_visibles["total"].mean()
        ganancia_mes = ventas_mes["ganancia"].sum()
    else:
        fact_hoy = fact_mes = fact_anio = ticket = ganancia_mes = 0

    perdida_total = perdidas_visibles["perdida_total"].sum() if not perdidas_visibles.empty else 0
    stock_valorizado = (stock_visible["stock"] * stock_visible["precio_compra"]).sum() if not stock_visible.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("Ventas de hoy", pesos(fact_hoy), "🛒", "Actualizado en tiempo real")
    with c2:
        metric_card("Ventas del mes", pesos(fact_mes), "📈", "Ciclo comercial actual")
    with c3:
        metric_card("Ganancia mensual", pesos(ganancia_mes), "💰", "Estimación bruta")
    with c4:
        metric_card("Ticket promedio", pesos(ticket), "🎟️", "Promedio por venta")

    c5, c6, c7, c8 = st.columns(4)
    with c5:
        metric_card("Ventas del año", pesos(fact_anio), "📅", "Acumulado anual")
    with c6:
        metric_card("Stock valorizado", pesos(stock_valorizado), "📦", "Costo de inventario")
    with c7:
        metric_card("Pérdidas registradas", pesos(perdida_total), "⚠️", "Control operativo")
    with c8:
        metric_card("Mensajes internos", len(chat_df), "💬", "Comunicación interna")

    st.divider()

    col1, col2 = st.columns([2, 1])

    with col1:
        st.subheader("📈 Evolución de ventas")
        if not ventas_visibles.empty:
            ventas_visibles["dia"] = ventas_visibles["fecha"].dt.date
            ventas_dia = ventas_visibles.groupby("dia")["total"].sum().reset_index()
            st.line_chart(ventas_dia.set_index("dia"))
        else:
            st.info("Todavía no hay ventas cargadas.")

    with col2:
        st.subheader("🏪 Ranking de sucursales")
        if not ventas_df.empty:
            ranking = ventas_df.groupby("sucursal")["total"].sum().reset_index()
            ranking = ranking.sort_values("total", ascending=False)
            st.dataframe(ranking, use_container_width=True, hide_index=True)
        else:
            st.info("Sin datos de ventas.")

    col3, col4, col5 = st.columns(3)

    with col3:
        st.subheader("🔥 Productos más vendidos")
        if not ventas_visibles.empty:
            top = ventas_visibles.groupby("producto")["cantidad"].sum().reset_index()
            top = top.sort_values("cantidad", ascending=False).head(8)
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

# =========================
# VENTAS
# =========================

elif menu == "Ventas":
    st.header("🛒 Gestión de ventas")

    if stock_visible.empty:
        st.warning("Primero cargá productos en stock.")
    else:
        with st.form("form_venta"):
            c1, c2, c3 = st.columns(3)

            if rol_actual == "Administrador":
                sucursal = c1.selectbox("Sucursal", SUCURSALES)
            else:
                sucursal = sucursal_usuario
                c1.text_input("Sucursal", value=sucursal, disabled=True)

            stock_sucursal = stock_df[stock_df["sucursal"] == sucursal]

            if stock_sucursal.empty:
                st.warning("No hay productos en esta sucursal.")
                producto_nombre = None
            else:
                producto_nombre = c2.selectbox("Producto", stock_sucursal["producto"].unique())

            cantidad = c3.number_input("Cantidad", min_value=1, step=1)

            c4, c5, c6 = st.columns(3)

            if producto_nombre:
                row = stock_sucursal[stock_sucursal["producto"] == producto_nombre].iloc[0]
                precio_unitario = c4.number_input(
                    "Precio unitario",
                    value=float(row["precio_venta"]),
                    min_value=0.0,
                    step=100.0
                )
                precio_compra = float(row["precio_compra"])
                stock_actual = int(row["stock"])
            else:
                row = None
                precio_unitario = c4.number_input("Precio unitario", min_value=0.0, step=100.0)
                precio_compra = 0
                stock_actual = 0

            descuento = c5.number_input("Descuento", min_value=0.0, value=0.0, step=100.0)
            medio_pago = c6.selectbox("Medio de pago", MEDIOS_PAGO)

            c7, c8, c9 = st.columns(3)
            vendedor = c7.text_input("Vendedor", value=usuario_actual)
            cliente = c8.text_input("Cliente")
            comprobante = c9.text_input("Comprobante")

            total = cantidad * precio_unitario - descuento
            costo_total = cantidad * precio_compra
            ganancia = total - costo_total
            margen = (ganancia / total * 100) if total > 0 else 0

            st.info(
                f"Total: {pesos(total)} | Ganancia: {pesos(ganancia)} | "
                f"Margen: {margen:.2f}% | Stock actual: {stock_actual}"
            )

            guardar = st.form_submit_button("Guardar venta")

            if guardar:
                if row is None:
                    st.error("Seleccioná un producto válido.")
                elif cantidad > stock_actual:
                    st.error("No hay stock suficiente.")
                else:
                    nueva = {
                        "fecha": str(datetime.now()),
                        "sucursal": sucursal,
                        "codigo": row["codigo"],
                        "producto": row["producto"],
                        "categoria": row["categoria"],
                        "cantidad": cantidad,
                        "precio_unitario": precio_unitario,
                        "descuento": descuento,
                        "total": total,
                        "costo_total": costo_total,
                        "ganancia": ganancia,
                        "medio_pago": medio_pago,
                        "vendedor": vendedor,
                        "cliente": cliente,
                        "comprobante": comprobante
                    }

                    ventas_df = pd.concat([ventas_df, pd.DataFrame([nueva])], ignore_index=True)
                    guardar_csv(ventas_df, VENTAS_FILE)

                    idx = stock_df[
                        (stock_df["codigo"] == row["codigo"]) &
                        (stock_df["sucursal"] == sucursal)
                    ].index[0]

                    stock_df.loc[idx, "stock"] = stock_actual - cantidad
                    guardar_csv(stock_df, STOCK_FILE)

                    st.success("Venta guardada correctamente.")
                    st.rerun()

    st.subheader("Historial de ventas")
    st.dataframe(ventas_visibles, use_container_width=True, hide_index=True)

# =========================
# STOCK
# =========================

elif menu == "Stock":
    st.header("📦 Gestión de stock")

    c1, c2, c3 = st.columns(3)

    if rol_actual == "Administrador":
        filtro_sucursal = c1.selectbox("Sucursal", ["Todas"] + SUCURSALES)
    else:
        filtro_sucursal = sucursal_usuario
        c1.text_input("Sucursal", value=sucursal_usuario, disabled=True)

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

    with st.form("form_stock"):
        c1, c2, c3 = st.columns(3)
        codigo = c1.text_input("Código")
        producto = c2.text_input("Producto")
        categoria = c3.selectbox("Categoría", CATEGORIAS)

        c4, c5, c6 = st.columns(3)
        marca = c4.text_input("Marca")
        proveedor = c5.text_input("Proveedor")

        if rol_actual == "Administrador":
            sucursal = c6.selectbox("Sucursal", SUCURSALES)
        else:
            sucursal = sucursal_usuario
            c6.text_input("Sucursal", value=sucursal, disabled=True)

        c7, c8, c9 = st.columns(3)
        precio_compra = c7.number_input("Precio compra", min_value=0.0, step=100.0)
        precio_venta = c8.number_input("Precio venta", min_value=0.0, step=100.0)
        cantidad = c9.number_input("Stock", min_value=0, step=1)

        c10, c11, c12 = st.columns(3)
        stock_minimo = c10.number_input("Stock mínimo", min_value=0, step=1)
        fecha_ingreso = c11.date_input("Fecha ingreso", value=date.today())
        fecha_vencimiento = c12.text_input("Fecha vencimiento")

        guardar = st.form_submit_button("Guardar producto")

        if guardar:
            nuevo = {
                "codigo": codigo,
                "producto": producto,
                "categoria": categoria,
                "marca": marca,
                "proveedor": proveedor,
                "sucursal": sucursal,
                "precio_compra": precio_compra,
                "precio_venta": precio_venta,
                "stock": cantidad,
                "stock_minimo": stock_minimo,
                "fecha_ingreso": str(fecha_ingreso),
                "fecha_vencimiento": fecha_vencimiento
            }

            stock_df = pd.concat([stock_df, pd.DataFrame([nuevo])], ignore_index=True)
            guardar_csv(stock_df, STOCK_FILE)
            st.success("Producto guardado.")
            st.rerun()

# =========================
# SUCURSALES
# =========================

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
    with c1:
        metric_card("Facturación", pesos(fact), "💰")
    with c2:
        metric_card("Ganancia", pesos(gan), "📈")
    with c3:
        metric_card("Pérdidas", pesos(per), "⚠️")
    with c4:
        metric_card("Inventario", pesos(inv), "📦")

    st.subheader("Ventas")
    st.dataframe(v, use_container_width=True, hide_index=True)

    st.subheader("Stock")
    st.dataframe(s, use_container_width=True, hide_index=True)

# =========================
# CLIENTES
# =========================

elif menu == "Clientes":
    st.header("👥 Clientes")

    with st.form("form_cliente"):
        c1, c2, c3 = st.columns(3)
        nombre = c1.text_input("Nombre")
        telefono = c2.text_input("Teléfono")
        email = c3.text_input("Email")

        c4, c5, c6 = st.columns(3)
        mascota = c4.text_input("Mascota")
        tipo = c5.selectbox("Tipo de mascota", ["Perro", "Gato", "Ave", "Roedor", "Otro"])
        obs = c6.text_input("Observaciones")

        guardar = st.form_submit_button("Guardar cliente")

        if guardar:
            nuevo = {
                "nombre": nombre,
                "telefono": telefono,
                "mascota": mascota,
                "tipo_mascota": tipo,
                "email": email,
                "observaciones": obs
            }
            clientes_df = pd.concat([clientes_df, pd.DataFrame([nuevo])], ignore_index=True)
            guardar_csv(clientes_df, CLIENTES_FILE)
            st.success("Cliente guardado.")
            st.rerun()

    st.dataframe(clientes_df, use_container_width=True, hide_index=True)

# =========================
# COMPRAS
# =========================

elif menu == "Compras":
    st.header("🚚 Compras a proveedores")

    with st.form("form_compra"):
        c1, c2, c3 = st.columns(3)
        proveedor = c1.text_input("Proveedor")
        codigo = c2.text_input("Código")
        producto = c3.text_input("Producto")

        c4, c5, c6 = st.columns(3)
        categoria = c4.selectbox("Categoría", CATEGORIAS)

        if rol_actual == "Administrador":
            sucursal = c5.selectbox("Sucursal destino", SUCURSALES)
        else:
            sucursal = sucursal_usuario
            c5.text_input("Sucursal destino", value=sucursal, disabled=True)

        cantidad = c6.number_input("Cantidad", min_value=1, step=1)

        c7, c8 = st.columns(2)
        precio_compra = c7.number_input("Precio compra unitario", min_value=0.0, step=100.0)
        precio_venta = c8.number_input("Precio venta sugerido", min_value=0.0, step=100.0)

        costo_total = cantidad * precio_compra
        st.info(f"Costo total: {pesos(costo_total)}")

        guardar = st.form_submit_button("Guardar compra")

        if guardar:
            nueva = {
                "fecha": str(datetime.now()),
                "proveedor": proveedor,
                "codigo": codigo,
                "producto": producto,
                "categoria": categoria,
                "sucursal": sucursal,
                "cantidad": cantidad,
                "precio_compra_unitario": precio_compra,
                "costo_total": costo_total
            }

            compras_df = pd.concat([compras_df, pd.DataFrame([nueva])], ignore_index=True)
            guardar_csv(compras_df, COMPRAS_FILE)

            existe = stock_df[
                (stock_df["codigo"] == codigo) &
                (stock_df["sucursal"] == sucursal)
            ]

            if not existe.empty:
                idx = existe.index[0]
                stock_df.loc[idx, "stock"] += cantidad
                stock_df.loc[idx, "precio_compra"] = precio_compra
                stock_df.loc[idx, "precio_venta"] = precio_venta
            else:
                nuevo_stock = {
                    "codigo": codigo,
                    "producto": producto,
                    "categoria": categoria,
                    "marca": "",
                    "proveedor": proveedor,
                    "sucursal": sucursal,
                    "precio_compra": precio_compra,
                    "precio_venta": precio_venta,
                    "stock": cantidad,
                    "stock_minimo": 5,
                    "fecha_ingreso": str(date.today()),
                    "fecha_vencimiento": ""
                }
                stock_df = pd.concat([stock_df, pd.DataFrame([nuevo_stock])], ignore_index=True)

            guardar_csv(stock_df, STOCK_FILE)
            st.success("Compra guardada y stock actualizado.")
            st.rerun()

    st.dataframe(compras_visibles, use_container_width=True, hide_index=True)

# =========================
# GANANCIAS Y PÉRDIDAS
# =========================

elif menu == "Ganancias y pérdidas":
    st.header("💰 Ganancias y pérdidas")

    total_fact = ventas_df["total"].sum() if not ventas_df.empty else 0
    total_costo = ventas_df["costo_total"].sum() if not ventas_df.empty else 0
    total_gan = ventas_df["ganancia"].sum() if not ventas_df.empty else 0
    total_perd = perdidas_df["perdida_total"].sum() if not perdidas_df.empty else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        metric_card("Facturación", pesos(total_fact), "💰")
    with c2:
        metric_card("Costo vendido", pesos(total_costo), "📦")
    with c3:
        metric_card("Ganancia bruta", pesos(total_gan), "📈")
    with c4:
        metric_card("Resultado neto", pesos(total_gan - total_perd), "🏆")

    st.subheader("Registrar pérdida")

    with st.form("form_perdida"):
        c1, c2, c3 = st.columns(3)
        sucursal = c1.selectbox("Sucursal", SUCURSALES)
        stock_sucursal = stock_df[stock_df["sucursal"] == sucursal]

        if not stock_sucursal.empty:
            producto = c2.selectbox("Producto", stock_sucursal["producto"].unique())
            row = stock_sucursal[stock_sucursal["producto"] == producto].iloc[0]
            codigo = row["codigo"]
            costo = float(row["precio_compra"])
        else:
            producto = c2.text_input("Producto")
            codigo = ""
            costo = 0

        motivo = c3.selectbox("Motivo", ["Vencimiento", "Rotura", "Faltante", "Descuento", "Otro"])

        c4, c5 = st.columns(2)
        cantidad = c4.number_input("Cantidad", min_value=1, step=1)
        obs = c5.text_input("Observaciones")

        perdida = cantidad * costo
        st.info(f"Pérdida estimada: {pesos(perdida)}")

        guardar = st.form_submit_button("Guardar pérdida")

        if guardar:
            nueva = {
                "fecha": str(datetime.now()),
                "sucursal": sucursal,
                "codigo": codigo,
                "producto": producto,
                "motivo": motivo,
                "cantidad": cantidad,
                "costo_unitario": costo,
                "perdida_total": perdida,
                "observaciones": obs
            }

            perdidas_df = pd.concat([perdidas_df, pd.DataFrame([nueva])], ignore_index=True)
            guardar_csv(perdidas_df, PERDIDAS_FILE)
            st.success("Pérdida registrada.")
            st.rerun()

    st.dataframe(perdidas_df, use_container_width=True, hide_index=True)

# =========================
# REPORTES
# =========================

elif menu == "Reportes":
    st.header("📊 Reportes avanzados")

    if ventas_visibles.empty:
        st.info("No hay ventas para analizar.")
    else:
        ventas_visibles["fecha"] = pd.to_datetime(ventas_visibles["fecha"], errors="coerce")
        ventas_visibles["dia"] = ventas_visibles["fecha"].dt.date
        ventas_visibles["mes"] = ventas_visibles["fecha"].dt.to_period("M").astype(str)

        st.subheader("Ventas por día")
        ventas_dia = ventas_visibles.groupby("dia")["total"].sum().reset_index()
        st.line_chart(ventas_dia.set_index("dia"))

        st.subheader("Ventas por mes")
        ventas_mes = ventas_visibles.groupby("mes")["total"].sum().reset_index()
        st.bar_chart(ventas_mes.set_index("mes"))

        st.subheader("Productos más rentables")
        rentables = ventas_visibles.groupby("producto")["ganancia"].sum().reset_index()
        rentables = rentables.sort_values("ganancia", ascending=False)
        st.dataframe(rentables, use_container_width=True, hide_index=True)

        st.subheader("Ventas por categoría")
        cat = ventas_visibles.groupby("categoria")["total"].sum().reset_index()
        st.bar_chart(cat.set_index("categoria"))

        st.subheader("Medios de pago")
        medios = ventas_visibles.groupby("medio_pago")["total"].sum().reset_index()
        st.dataframe(medios, use_container_width=True, hide_index=True)

# =========================
# CHAT INTERNO
# =========================

elif menu == "Chat interno":
    st.header("💬 Chat interno")

    st.info(
        "Chat interno para comunicación entre sucursales, dueño, encargados y vendedores. "
        "Permite adjuntar imágenes, facturas, PDF, Excel y documentos."
    )

    chat_df = cargar_csv(CHAT_FILE)

    tipo_chat = st.radio(
        "Seleccionar tipo de chat",
        ["Chat general", "Chat por sucursal", "Chat privado"],
        horizontal=True
    )

    canal = "general"
    para_usuario = ""

    if tipo_chat == "Chat general":
        st.subheader("🌐 Chat general de la empresa")
        canal = "general"

    elif tipo_chat == "Chat por sucursal":
        if rol_actual == "Administrador":
            canal = st.selectbox("Sucursal", SUCURSALES)
        else:
            canal = sucursal_usuario
            st.text_input("Sucursal", value=canal, disabled=True)

        st.subheader(f"🏪 Chat de {canal}")

    elif tipo_chat == "Chat privado":
        usuarios_disponibles = [u for u in USUARIOS.keys() if u != usuario_actual]

        if rol_actual != "Administrador":
            usuarios_disponibles = ["admin"]

        para_usuario = st.selectbox("Hablar con", usuarios_disponibles)

        participantes = sorted([usuario_actual, para_usuario])
        canal = f"privado_{participantes[0]}_{participantes[1]}"

        st.subheader(f"🔒 Chat privado con {para_usuario}")

    st.divider()

    with st.form("form_chat", clear_on_submit=True):
        mensaje = st.text_area("Escribir mensaje", height=90)
        archivo = st.file_uploader(
            "Adjuntar archivo",
            type=["png", "jpg", "jpeg", "pdf", "xlsx", "xls", "docx", "txt"]
        )

        enviar = st.form_submit_button("Enviar mensaje")

        if enviar:
            if not mensaje and archivo is None:
                st.warning("Escribí un mensaje o adjuntá un archivo.")
            else:
                archivo_nombre, archivo_ruta, archivo_tipo = guardar_adjunto(archivo, usuario_actual)

                nuevo_mensaje = {
                    "fecha": str(datetime.now()),
                    "tipo_chat": tipo_chat,
                    "canal": canal,
                    "de_usuario": usuario_actual,
                    "de_nombre": nombre_actual,
                    "para_usuario": para_usuario,
                    "mensaje": mensaje,
                    "archivo_nombre": archivo_nombre,
                    "archivo_ruta": archivo_ruta,
                    "archivo_tipo": archivo_tipo
                }

                chat_df = pd.concat([chat_df, pd.DataFrame([nuevo_mensaje])], ignore_index=True)
                guardar_csv(chat_df, CHAT_FILE)
                st.success("Mensaje enviado.")
                st.rerun()

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

                st.markdown(
                    f"""
                    <div class="chat-message">
                        <div class="chat-head">👤 {msg['de_nombre']} · {msg['de_usuario']} · {fecha_txt}</div>
                        <div class="chat-text">{texto}</div>
                    </div>
                    """,
                    unsafe_allow_html=True
                )

                if pd.notnull(msg.get("archivo_ruta", "")) and str(msg.get("archivo_ruta", "")) != "":
                    mostrar_archivo_chat(
                        msg.get("archivo_nombre", ""),
                        msg.get("archivo_ruta", ""),
                        msg.get("archivo_tipo", "")
                    )
        else:
            st.info("No hay mensajes en este chat todavía.")

    if rol_actual == "Administrador":
        st.divider()
        st.subheader("📋 Auditoría de mensajes")
        st.dataframe(chat_df, use_container_width=True, hide_index=True)

# =========================
# IMPORTAR EXCEL
# =========================

elif menu == "Importar Excel":
    st.header("📥 Importar Excel")

    tipo = st.selectbox("Tipo de datos", ["Stock", "Ventas", "Compras", "Clientes"])
    archivo = st.file_uploader("Subir Excel", type=["xlsx", "xls"])

    if archivo:
        df_importado = pd.read_excel(archivo)
        st.dataframe(df_importado, use_container_width=True, hide_index=True)

        if st.button("Importar y reemplazar datos"):
            if tipo == "Stock":
                guardar_csv(df_importado, STOCK_FILE)
            elif tipo == "Ventas":
                guardar_csv(df_importado, VENTAS_FILE)
            elif tipo == "Compras":
                guardar_csv(df_importado, COMPRAS_FILE)
            elif tipo == "Clientes":
                guardar_csv(df_importado, CLIENTES_FILE)

            st.success("Datos importados.")
            st.rerun()

# =========================
# EXPORTAR
# =========================

elif menu == "Exportar datos":
    st.header("📤 Exportar datos")

    excel_data = exportar_excel({
        "Ventas": ventas_df,
        "Stock": stock_df,
        "Compras": compras_df,
        "Clientes": clientes_df,
        "Perdidas": perdidas_df,
        "Chat": chat_df
    })

    st.download_button(
        "Descargar Excel completo",
        data=excel_data,
        file_name="petshop_manager_pro.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

# =========================
# CONFIGURACIÓN
# =========================

elif menu == "Configuración":
    st.header("⚙️ Configuración del sistema")

    st.info("Usuarios demo disponibles:")

    usuarios_demo = pd.DataFrame([
        {"Usuario": "admin", "Contraseña": "1234", "Rol": "Administrador", "Acceso": "Todo el sistema"},
        {"Usuario": "encargado", "Contraseña": "1234", "Rol": "Encargado", "Acceso": "Sucursal asignada"},
        {"Usuario": "vendedor", "Contraseña": "1234", "Rol": "Vendedor", "Acceso": "Ventas, clientes y chat"},
        {"Usuario": "norte", "Contraseña": "1234", "Rol": "Vendedor", "Acceso": "Sucursal Norte"},
        {"Usuario": "sur", "Contraseña": "1234", "Rol": "Vendedor", "Acceso": "Sucursal Sur"},
    ])

    st.dataframe(usuarios_demo, use_container_width=True, hide_index=True)

    st.warning(
        "Esta versión guarda datos, chat y adjuntos en archivos locales. "
        "Para uso real entre varias sucursales online, el próximo paso es Supabase/Firebase."
    )
