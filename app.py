import streamlit as st
import pandas as pd
from datetime import datetime, date
import os
from io import BytesIO

st.set_page_config(
    page_title="Pet Shop Manager",
    page_icon="🐾",
    layout="wide"
)

DATA_DIR = "data_petshop"
VENTAS_FILE = f"{DATA_DIR}/ventas.csv"
STOCK_FILE = f"{DATA_DIR}/stock.csv"
COMPRAS_FILE = f"{DATA_DIR}/compras.csv"
CLIENTES_FILE = f"{DATA_DIR}/clientes.csv"
PERDIDAS_FILE = f"{DATA_DIR}/perdidas.csv"

SUCURSALES = ["Sucursal Centro", "Sucursal Norte", "Sucursal Sur"]

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

MEDIOS_PAGO = [
    "Efectivo",
    "Débito",
    "Crédito",
    "Transferencia",
    "Mercado Pago"
]

os.makedirs(DATA_DIR, exist_ok=True)


def crear_archivos_iniciales():
    if not os.path.exists(STOCK_FILE):
        stock = pd.DataFrame([
            {
                "codigo": "DOG001",
                "producto": "Alimento Perro Adulto 15kg",
                "categoria": "Alimento para perros",
                "marca": "DogPlus",
                "proveedor": "Proveedor A",
                "sucursal": "Sucursal Centro",
                "precio_compra": 18000,
                "precio_venta": 26000,
                "stock": 25,
                "stock_minimo": 5,
                "fecha_ingreso": str(date.today()),
                "fecha_vencimiento": ""
            },
            {
                "codigo": "CAT001",
                "producto": "Alimento Gato Adulto 10kg",
                "categoria": "Alimento para gatos",
                "marca": "CatPremium",
                "proveedor": "Proveedor B",
                "sucursal": "Sucursal Norte",
                "precio_compra": 14000,
                "precio_venta": 21000,
                "stock": 18,
                "stock_minimo": 4,
                "fecha_ingreso": str(date.today()),
                "fecha_vencimiento": ""
            },
            {
                "codigo": "ARE001",
                "producto": "Arena Sanitaria 4kg",
                "categoria": "Arena para gatos",
                "marca": "MichiClean",
                "proveedor": "Proveedor C",
                "sucursal": "Sucursal Sur",
                "precio_compra": 2500,
                "precio_venta": 4200,
                "stock": 40,
                "stock_minimo": 10,
                "fecha_ingreso": str(date.today()),
                "fecha_vencimiento": ""
            }
        ])
        stock.to_csv(STOCK_FILE, index=False)

    if not os.path.exists(VENTAS_FILE):
        ventas = pd.DataFrame(columns=[
            "fecha",
            "sucursal",
            "codigo",
            "producto",
            "categoria",
            "cantidad",
            "precio_unitario",
            "descuento",
            "total",
            "costo_total",
            "ganancia",
            "medio_pago",
            "vendedor",
            "cliente",
            "comprobante"
        ])
        ventas.to_csv(VENTAS_FILE, index=False)

    if not os.path.exists(COMPRAS_FILE):
        compras = pd.DataFrame(columns=[
            "fecha",
            "proveedor",
            "codigo",
            "producto",
            "categoria",
            "sucursal",
            "cantidad",
            "precio_compra_unitario",
            "costo_total"
        ])
        compras.to_csv(COMPRAS_FILE, index=False)

    if not os.path.exists(CLIENTES_FILE):
        clientes = pd.DataFrame(columns=[
            "nombre",
            "telefono",
            "mascota",
            "tipo_mascota",
            "email",
            "observaciones"
        ])
        clientes.to_csv(CLIENTES_FILE, index=False)

    if not os.path.exists(PERDIDAS_FILE):
        perdidas = pd.DataFrame(columns=[
            "fecha",
            "sucursal",
            "codigo",
            "producto",
            "motivo",
            "cantidad",
            "costo_unitario",
            "perdida_total",
            "observaciones"
        ])
        perdidas.to_csv(PERDIDAS_FILE, index=False)


def cargar_csv(path):
    if os.path.exists(path):
        return pd.read_csv(path)
    return pd.DataFrame()


def guardar_csv(df, path):
    df.to_csv(path, index=False)


def formato_pesos(valor):
    try:
        return f"${valor:,.0f}".replace(",", ".")
    except:
        return "$0"


def exportar_excel(dfs):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for nombre, df in dfs.items():
            df.to_excel(writer, index=False, sheet_name=nombre[:31])
    return output.getvalue()


crear_archivos_iniciales()

stock_df = cargar_csv(STOCK_FILE)
ventas_df = cargar_csv(VENTAS_FILE)
compras_df = cargar_csv(COMPRAS_FILE)
clientes_df = cargar_csv(CLIENTES_FILE)
perdidas_df = cargar_csv(PERDIDAS_FILE)

st.markdown("""
<style>
.main {
    background-color: #f7fbfa;
}
.block-container {
    padding-top: 1.5rem;
}
.metric-card {
    background: white;
    padding: 18px;
    border-radius: 18px;
    border: 1px solid #e6eeee;
    box-shadow: 0px 4px 18px rgba(0,0,0,0.04);
}
h1, h2, h3 {
    color: #164e42;
}
.stButton>button {
    background-color: #20b486;
    color: white;
    border-radius: 10px;
    border: none;
    padding: 0.6rem 1rem;
}
.stButton>button:hover {
    background-color: #14956d;
    color: white;
}
</style>
""", unsafe_allow_html=True)

st.sidebar.title("🐾 Pet Shop Manager")
st.sidebar.caption("Sistema profesional para 3 sucursales")

menu = st.sidebar.radio(
    "Menú principal",
    [
        "Dashboard",
        "Ventas",
        "Stock",
        "Sucursales",
        "Compras a proveedores",
        "Clientes",
        "Ganancias y pérdidas",
        "Reportes",
        "Importar Excel",
        "Exportar datos"
    ]
)

st.title("🐾 Pet Shop Manager")
st.caption("Control de ventas, stock, facturación, ganancias y sucursales")


if menu == "Dashboard":
    st.header("📊 Dashboard principal")

    if not ventas_df.empty:
        ventas_df["fecha"] = pd.to_datetime(ventas_df["fecha"], errors="coerce")
        hoy = pd.to_datetime(date.today())

        ventas_hoy = ventas_df[ventas_df["fecha"].dt.date == date.today()]
        ventas_mes = ventas_df[
            (ventas_df["fecha"].dt.month == hoy.month) &
            (ventas_df["fecha"].dt.year == hoy.year)
        ]
        ventas_anio = ventas_df[ventas_df["fecha"].dt.year == hoy.year]

        facturacion_hoy = ventas_hoy["total"].sum()
        facturacion_mes = ventas_mes["total"].sum()
        facturacion_anio = ventas_anio["total"].sum()
        ganancia_mes = ventas_mes["ganancia"].sum()
        cantidad_ventas = len(ventas_df)
        ticket_promedio = ventas_df["total"].mean() if len(ventas_df) > 0 else 0
    else:
        facturacion_hoy = 0
        facturacion_mes = 0
        facturacion_anio = 0
        ganancia_mes = 0
        cantidad_ventas = 0
        ticket_promedio = 0

    perdida_total = perdidas_df["perdida_total"].sum() if not perdidas_df.empty else 0

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Facturación hoy", formato_pesos(facturacion_hoy))
    col2.metric("Facturación mensual", formato_pesos(facturacion_mes))
    col3.metric("Ganancia mensual", formato_pesos(ganancia_mes))
    col4.metric("Pérdidas", formato_pesos(perdida_total))

    col5, col6, col7, col8 = st.columns(4)
    col5.metric("Facturación anual", formato_pesos(facturacion_anio))
    col6.metric("Cantidad de ventas", cantidad_ventas)
    col7.metric("Ticket promedio", formato_pesos(ticket_promedio))
    col8.metric("Stock valorizado", formato_pesos((stock_df["stock"] * stock_df["precio_compra"]).sum()))

    st.divider()

    col_a, col_b = st.columns(2)

    with col_a:
        st.subheader("🏆 Ranking de sucursales")
        if not ventas_df.empty:
            ranking = ventas_df.groupby("sucursal")["total"].sum().reset_index()
            ranking = ranking.sort_values("total", ascending=False)
            st.bar_chart(ranking.set_index("sucursal"))
            st.dataframe(ranking, use_container_width=True)
        else:
            st.info("Todavía no hay ventas cargadas.")

    with col_b:
        st.subheader("⚠️ Productos con bajo stock")
        bajo_stock = stock_df[stock_df["stock"] <= stock_df["stock_minimo"]]
        if not bajo_stock.empty:
            st.warning("Hay productos que necesitan reposición.")
            st.dataframe(bajo_stock, use_container_width=True)
        else:
            st.success("No hay productos con bajo stock.")

    st.subheader("🔥 Productos más vendidos")
    if not ventas_df.empty:
        productos = ventas_df.groupby("producto")["cantidad"].sum().reset_index()
        productos = productos.sort_values("cantidad", ascending=False).head(10)
        st.bar_chart(productos.set_index("producto"))
        st.dataframe(productos, use_container_width=True)
    else:
        st.info("Todavía no hay productos vendidos.")


elif menu == "Ventas":
    st.header("🛒 Cargar venta diaria")

    if stock_df.empty:
        st.warning("Primero cargá productos en stock.")
    else:
        with st.form("form_venta"):
            col1, col2, col3 = st.columns(3)

            sucursal = col1.selectbox("Sucursal", SUCURSALES)
            stock_sucursal = stock_df[stock_df["sucursal"] == sucursal]

            if stock_sucursal.empty:
                st.warning("No hay productos cargados para esta sucursal.")
                producto_seleccionado = None
            else:
                producto_nombre = col2.selectbox("Producto", stock_sucursal["producto"].unique())
                producto_row = stock_sucursal[stock_sucursal["producto"] == producto_nombre].iloc[0]
                producto_seleccionado = producto_row

            cantidad = col3.number_input("Cantidad vendida", min_value=1, step=1)

            col4, col5, col6 = st.columns(3)

            if producto_seleccionado is not None:
                precio_unitario = col4.number_input(
                    "Precio unitario",
                    min_value=0.0,
                    value=float(producto_seleccionado["precio_venta"]),
                    step=100.0
                )
                precio_compra = float(producto_seleccionado["precio_compra"])
                stock_actual = int(producto_seleccionado["stock"])
            else:
                precio_unitario = col4.number_input("Precio unitario", min_value=0.0, step=100.0)
                precio_compra = 0
                stock_actual = 0

            descuento = col5.number_input("Descuento", min_value=0.0, value=0.0, step=100.0)
            medio_pago = col6.selectbox("Medio de pago", MEDIOS_PAGO)

            col7, col8, col9 = st.columns(3)
            vendedor = col7.text_input("Vendedor")
            cliente = col8.text_input("Cliente")
            comprobante = col9.text_input("Comprobante / Nº venta")

            total = cantidad * precio_unitario - descuento
            costo_total = cantidad * precio_compra
            ganancia = total - costo_total
            margen = (ganancia / total * 100) if total > 0 else 0

            st.info(
                f"Total venta: {formato_pesos(total)} | "
                f"Ganancia estimada: {formato_pesos(ganancia)} | "
                f"Margen: {margen:.2f}% | "
                f"Stock actual: {stock_actual}"
            )

            guardar = st.form_submit_button("Guardar venta")

            if guardar:
                if producto_seleccionado is None:
                    st.error("No seleccionaste un producto válido.")
                elif cantidad > stock_actual:
                    st.error("No hay stock suficiente para esta venta.")
                else:
                    nueva_venta = {
                        "fecha": str(datetime.now()),
                        "sucursal": sucursal,
                        "codigo": producto_seleccionado["codigo"],
                        "producto": producto_seleccionado["producto"],
                        "categoria": producto_seleccionado["categoria"],
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

                    ventas_df = pd.concat([ventas_df, pd.DataFrame([nueva_venta])], ignore_index=True)
                    guardar_csv(ventas_df, VENTAS_FILE)

                    idx = stock_df[
                        (stock_df["codigo"] == producto_seleccionado["codigo"]) &
                        (stock_df["sucursal"] == sucursal)
                    ].index[0]

                    stock_df.loc[idx, "stock"] = stock_actual - cantidad
                    guardar_csv(stock_df, STOCK_FILE)

                    st.success("Venta guardada correctamente.")
                    st.rerun()

    st.subheader("Historial de ventas")
    st.dataframe(ventas_df, use_container_width=True)


elif menu == "Stock":
    st.header("📦 Stock general y por sucursal")

    col1, col2, col3 = st.columns(3)
    filtro_sucursal = col1.selectbox("Filtrar por sucursal", ["Todas"] + SUCURSALES)
    filtro_categoria = col2.selectbox("Filtrar por categoría", ["Todas"] + CATEGORIAS)
    buscar = col3.text_input("Buscar producto")

    stock_filtrado = stock_df.copy()

    if filtro_sucursal != "Todas":
        stock_filtrado = stock_filtrado[stock_filtrado["sucursal"] == filtro_sucursal]

    if filtro_categoria != "Todas":
        stock_filtrado = stock_filtrado[stock_filtrado["categoria"] == filtro_categoria]

    if buscar:
        stock_filtrado = stock_filtrado[
            stock_filtrado["producto"].str.contains(buscar, case=False, na=False)
        ]

    st.dataframe(stock_filtrado, use_container_width=True)

    st.divider()
    st.subheader("➕ Agregar producto al stock")

    with st.form("form_stock"):
        col1, col2, col3 = st.columns(3)
        codigo = col1.text_input("Código")
        producto = col2.text_input("Producto")
        categoria = col3.selectbox("Categoría", CATEGORIAS)

        col4, col5, col6 = st.columns(3)
        marca = col4.text_input("Marca")
        proveedor = col5.text_input("Proveedor")
        sucursal = col6.selectbox("Sucursal", SUCURSALES)

        col7, col8, col9 = st.columns(3)
        precio_compra = col7.number_input("Precio de compra", min_value=0.0, step=100.0)
        precio_venta = col8.number_input("Precio de venta", min_value=0.0, step=100.0)
        cantidad = col9.number_input("Cantidad disponible", min_value=0, step=1)

        col10, col11, col12 = st.columns(3)
        stock_minimo = col10.number_input("Stock mínimo", min_value=0, step=1)
        fecha_ingreso = col11.date_input("Fecha de ingreso", value=date.today())
        fecha_vencimiento = col12.text_input("Fecha de vencimiento si aplica")

        guardar_producto = st.form_submit_button("Guardar producto")

        if guardar_producto:
            nuevo_producto = {
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

            stock_df = pd.concat([stock_df, pd.DataFrame([nuevo_producto])], ignore_index=True)
            guardar_csv(stock_df, STOCK_FILE)
            st.success("Producto agregado correctamente.")
            st.rerun()

    st.divider()
    st.subheader("🔁 Transferir stock entre sucursales")

    with st.form("form_transferencia"):
        col1, col2, col3 = st.columns(3)
        origen = col1.selectbox("Sucursal origen", SUCURSALES)
        destino = col2.selectbox("Sucursal destino", SUCURSALES)

        stock_origen = stock_df[stock_df["sucursal"] == origen]

        if not stock_origen.empty:
            producto_transferir = col3.selectbox("Producto a transferir", stock_origen["producto"].unique())
            cantidad_transferir = st.number_input("Cantidad a transferir", min_value=1, step=1)
        else:
            producto_transferir = None
            cantidad_transferir = 0
            st.warning("La sucursal origen no tiene stock.")

        transferir = st.form_submit_button("Transferir")

        if transferir:
            if origen == destino:
                st.error("La sucursal origen y destino no pueden ser iguales.")
            elif producto_transferir is None:
                st.error("No hay producto para transferir.")
            else:
                row_origen = stock_df[
                    (stock_df["sucursal"] == origen) &
                    (stock_df["producto"] == producto_transferir)
                ].iloc[0]

                idx_origen = row_origen.name
                stock_actual = int(row_origen["stock"])

                if cantidad_transferir > stock_actual:
                    st.error("No hay stock suficiente para transferir.")
                else:
                    stock_df.loc[idx_origen, "stock"] = stock_actual - cantidad_transferir

                    existe_destino = stock_df[
                        (stock_df["sucursal"] == destino) &
                        (stock_df["codigo"] == row_origen["codigo"])
                    ]

                    if not existe_destino.empty:
                        idx_destino = existe_destino.index[0]
                        stock_df.loc[idx_destino, "stock"] += cantidad_transferir
                    else:
                        nuevo = row_origen.copy()
                        nuevo["sucursal"] = destino
                        nuevo["stock"] = cantidad_transferir
                        stock_df = pd.concat([stock_df, pd.DataFrame([nuevo])], ignore_index=True)

                    guardar_csv(stock_df, STOCK_FILE)
                    st.success("Transferencia realizada correctamente.")
                    st.rerun()


elif menu == "Sucursales":
    st.header("🏪 Análisis por sucursal")

    sucursal = st.selectbox("Seleccionar sucursal", SUCURSALES)

    ventas_sucursal = ventas_df[ventas_df["sucursal"] == sucursal] if not ventas_df.empty else pd.DataFrame()
    stock_sucursal = stock_df[stock_df["sucursal"] == sucursal]
    perdidas_sucursal = perdidas_df[perdidas_df["sucursal"] == sucursal] if not perdidas_df.empty else pd.DataFrame()

    facturacion = ventas_sucursal["total"].sum() if not ventas_sucursal.empty else 0
    ganancia = ventas_sucursal["ganancia"].sum() if not ventas_sucursal.empty else 0
    perdida = perdidas_sucursal["perdida_total"].sum() if not perdidas_sucursal.empty else 0
    stock_valorizado = (stock_sucursal["stock"] * stock_sucursal["precio_compra"]).sum()

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Facturación", formato_pesos(facturacion))
    col2.metric("Ganancia", formato_pesos(ganancia))
    col3.metric("Pérdidas", formato_pesos(perdida))
    col4.metric("Stock valorizado", formato_pesos(stock_valorizado))

    st.subheader("Ventas de la sucursal")
    st.dataframe(ventas_sucursal, use_container_width=True)

    st.subheader("Stock de la sucursal")
    st.dataframe(stock_sucursal, use_container_width=True)


elif menu == "Compras a proveedores":
    st.header("🚚 Compras a proveedores")

    with st.form("form_compra"):
        col1, col2, col3 = st.columns(3)
        proveedor = col1.text_input("Proveedor")
        codigo = col2.text_input("Código producto")
        producto = col3.text_input("Producto")

        col4, col5, col6 = st.columns(3)
        categoria = col4.selectbox("Categoría", CATEGORIAS)
        sucursal = col5.selectbox("Sucursal destino", SUCURSALES)
        cantidad = col6.number_input("Cantidad comprada", min_value=1, step=1)

        col7, col8 = st.columns(2)
        precio_compra_unitario = col7.number_input("Precio compra unitario", min_value=0.0, step=100.0)
        precio_venta = col8.number_input("Precio venta sugerido", min_value=0.0, step=100.0)

        costo_total = cantidad * precio_compra_unitario
        st.info(f"Costo total de compra: {formato_pesos(costo_total)}")

        guardar_compra = st.form_submit_button("Guardar compra y actualizar stock")

        if guardar_compra:
            nueva_compra = {
                "fecha": str(datetime.now()),
                "proveedor": proveedor,
                "codigo": codigo,
                "producto": producto,
                "categoria": categoria,
                "sucursal": sucursal,
                "cantidad": cantidad,
                "precio_compra_unitario": precio_compra_unitario,
                "costo_total": costo_total
            }

            compras_df = pd.concat([compras_df, pd.DataFrame([nueva_compra])], ignore_index=True)
            guardar_csv(compras_df, COMPRAS_FILE)

            existe = stock_df[
                (stock_df["codigo"] == codigo) &
                (stock_df["sucursal"] == sucursal)
            ]

            if not existe.empty:
                idx = existe.index[0]
                stock_df.loc[idx, "stock"] += cantidad
                stock_df.loc[idx, "precio_compra"] = precio_compra_unitario
                stock_df.loc[idx, "precio_venta"] = precio_venta
            else:
                nuevo_stock = {
                    "codigo": codigo,
                    "producto": producto,
                    "categoria": categoria,
                    "marca": "",
                    "proveedor": proveedor,
                    "sucursal": sucursal,
                    "precio_compra": precio_compra_unitario,
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

    st.subheader("Historial de compras")
    st.dataframe(compras_df, use_container_width=True)


elif menu == "Clientes":
    st.header("👥 Clientes")

    with st.form("form_cliente"):
        col1, col2, col3 = st.columns(3)
        nombre = col1.text_input("Nombre del cliente")
        telefono = col2.text_input("Teléfono")
        email = col3.text_input("Email")

        col4, col5, col6 = st.columns(3)
        mascota = col4.text_input("Nombre de mascota")
        tipo_mascota = col5.selectbox("Tipo de mascota", ["Perro", "Gato", "Ave", "Roedor", "Otro"])
        observaciones = col6.text_input("Observaciones")

        guardar_cliente = st.form_submit_button("Guardar cliente")

        if guardar_cliente:
            nuevo_cliente = {
                "nombre": nombre,
                "telefono": telefono,
                "mascota": mascota,
                "tipo_mascota": tipo_mascota,
                "email": email,
                "observaciones": observaciones
            }

            clientes_df = pd.concat([clientes_df, pd.DataFrame([nuevo_cliente])], ignore_index=True)
            guardar_csv(clientes_df, CLIENTES_FILE)
            st.success("Cliente guardado correctamente.")
            st.rerun()

    st.subheader("Base de clientes")
    st.dataframe(clientes_df, use_container_width=True)

    st.subheader("Historial por cliente")
    if not ventas_df.empty:
        cliente_buscar = st.text_input("Buscar cliente en ventas")
        if cliente_buscar:
            historial = ventas_df[
                ventas_df["cliente"].str.contains(cliente_buscar, case=False, na=False)
            ]
            st.dataframe(historial, use_container_width=True)

            if not historial.empty:
                st.metric("Total gastado", formato_pesos(historial["total"].sum()))
                st.metric("Última compra", historial["fecha"].max())


elif menu == "Ganancias y pérdidas":
    st.header("💰 Ganancias y pérdidas")

    if not ventas_df.empty:
        total_facturado = ventas_df["total"].sum()
        costo_total = ventas_df["costo_total"].sum()
        ganancia_total = ventas_df["ganancia"].sum()
    else:
        total_facturado = 0
        costo_total = 0
        ganancia_total = 0

    perdida_total = perdidas_df["perdida_total"].sum() if not perdidas_df.empty else 0
    resultado_neto = ganancia_total - perdida_total

    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Facturación total", formato_pesos(total_facturado))
    col2.metric("Costo total vendido", formato_pesos(costo_total))
    col3.metric("Ganancia bruta", formato_pesos(ganancia_total))
    col4.metric("Resultado neto", formato_pesos(resultado_neto))

    st.divider()
    st.subheader("Registrar pérdida")

    with st.form("form_perdida"):
        col1, col2, col3 = st.columns(3)
        sucursal = col1.selectbox("Sucursal", SUCURSALES)
        stock_sucursal = stock_df[stock_df["sucursal"] == sucursal]

        if not stock_sucursal.empty:
            producto = col2.selectbox("Producto", stock_sucursal["producto"].unique())
            row_producto = stock_sucursal[stock_sucursal["producto"] == producto].iloc[0]
            codigo = row_producto["codigo"]
            costo_unitario = float(row_producto["precio_compra"])
            stock_actual = int(row_producto["stock"])
        else:
            producto = col2.text_input("Producto")
            codigo = ""
            costo_unitario = 0
            stock_actual = 0

        motivo = col3.selectbox(
            "Motivo",
            ["Vencimiento", "Rotura", "Faltante de stock", "Descuento", "Otro"]
        )

        col4, col5 = st.columns(2)
        cantidad = col4.number_input("Cantidad perdida", min_value=1, step=1)
        observaciones = col5.text_input("Observaciones")

        perdida_total_item = cantidad * costo_unitario
        st.info(f"Pérdida estimada: {formato_pesos(perdida_total_item)}")

        guardar_perdida = st.form_submit_button("Guardar pérdida")

        if guardar_perdida:
            nueva_perdida = {
                "fecha": str(datetime.now()),
                "sucursal": sucursal,
                "codigo": codigo,
                "producto": producto,
                "motivo": motivo,
                "cantidad": cantidad,
                "costo_unitario": costo_unitario,
                "perdida_total": perdida_total_item,
                "observaciones": observaciones
            }

            perdidas_df = pd.concat([perdidas_df, pd.DataFrame([nueva_perdida])], ignore_index=True)
            guardar_csv(perdidas_df, PERDIDAS_FILE)

            if not stock_sucursal.empty and cantidad <= stock_actual:
                idx = stock_df[
                    (stock_df["sucursal"] == sucursal) &
                    (stock_df["codigo"] == codigo)
                ].index[0]
                stock_df.loc[idx, "stock"] = stock_actual - cantidad
                guardar_csv(stock_df, STOCK_FILE)

            st.success("Pérdida registrada.")
            st.rerun()

    st.subheader("Historial de pérdidas")
    st.dataframe(perdidas_df, use_container_width=True)


elif menu == "Reportes":
    st.header("📈 Reportes")

    if ventas_df.empty:
        st.info("Todavía no hay ventas para generar reportes.")
    else:
        ventas_df["fecha"] = pd.to_datetime(ventas_df["fecha"], errors="coerce")
        ventas_df["mes"] = ventas_df["fecha"].dt.to_period("M").astype(str)
        ventas_df["dia"] = ventas_df["fecha"].dt.date

        st.subheader("Ventas por día")
        ventas_dia = ventas_df.groupby("dia")["total"].sum().reset_index()
        st.line_chart(ventas_dia.set_index("dia"))
        st.dataframe(ventas_dia, use_container_width=True)

        st.subheader("Ventas por mes")
        ventas_mes = ventas_df.groupby("mes")["total"].sum().reset_index()
        st.bar_chart(ventas_mes.set_index("mes"))
        st.dataframe(ventas_mes, use_container_width=True)

        st.subheader("Productos más vendidos")
        mas_vendidos = ventas_df.groupby("producto")["cantidad"].sum().reset_index()
        mas_vendidos = mas_vendidos.sort_values("cantidad", ascending=False)
        st.dataframe(mas_vendidos, use_container_width=True)

        st.subheader("Productos con mayor ganancia")
        mayor_ganancia = ventas_df.groupby("producto")["ganancia"].sum().reset_index()
        mayor_ganancia = mayor_ganancia.sort_values("ganancia", ascending=False)
        st.dataframe(mayor_ganancia, use_container_width=True)

        st.subheader("Facturación por categoría")
        categoria = ventas_df.groupby("categoria")["total"].sum().reset_index()
        st.bar_chart(categoria.set_index("categoria"))
        st.dataframe(categoria, use_container_width=True)

        st.subheader("Facturación por medio de pago")
        medio = ventas_df.groupby("medio_pago")["total"].sum().reset_index()
        st.dataframe(medio, use_container_width=True)


elif menu == "Importar Excel":
    st.header("📥 Importar datos desde Excel")

    st.info("Podés importar archivos Excel con columnas similares a las tablas del sistema.")

    tipo = st.selectbox("¿Qué querés importar?", ["Stock", "Ventas", "Compras", "Clientes"])

    archivo = st.file_uploader("Subir archivo Excel", type=["xlsx", "xls"])

    if archivo is not None:
        df_importado = pd.read_excel(archivo)
        st.subheader("Vista previa")
        st.dataframe(df_importado, use_container_width=True)

        if st.button("Importar y reemplazar datos actuales"):
            if tipo == "Stock":
                guardar_csv(df_importado, STOCK_FILE)
            elif tipo == "Ventas":
                guardar_csv(df_importado, VENTAS_FILE)
            elif tipo == "Compras":
                guardar_csv(df_importado, COMPRAS_FILE)
            elif tipo == "Clientes":
                guardar_csv(df_importado, CLIENTES_FILE)

            st.success("Datos importados correctamente.")
            st.rerun()


elif menu == "Exportar datos":
    st.header("📤 Exportar información")

    excel_data = exportar_excel({
        "Ventas": ventas_df,
        "Stock": stock_df,
        "Compras": compras_df,
        "Clientes": clientes_df,
        "Perdidas": perdidas_df
    })

    st.download_button(
        label="Descargar todo en Excel",
        data=excel_data,
        file_name="reporte_petshop.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.subheader("Datos disponibles")
    st.write("Ventas")
    st.dataframe(ventas_df, use_container_width=True)

    st.write("Stock")
    st.dataframe(stock_df, use_container_width=True)

    st.write("Compras")
    st.dataframe(compras_df, use_container_width=True)

    st.write("Clientes")
    st.dataframe(clientes_df, use_container_width=True)

    st.write("Pérdidas")
    st.dataframe(perdidas_df, use_container_width=True)