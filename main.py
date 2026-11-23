import io
import math
import os
import re
import unicodedata
from datetime import date, datetime, timedelta

import pandas as pd
import plotly.express as px
import plotly.io as pio
from flask import (
    Flask,
    flash,
    redirect,
    render_template,
    request,
    send_file,
    url_for,
)
from werkzeug.utils import secure_filename

from models import Producto

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 20 * 1024 * 1024
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY") or os.urandom(24)

CARPETA_EXCEL = os.environ.get("CARPETA_EXCEL")
if not CARPETA_EXCEL:
    if os.name == "nt":
        CARPETA_EXCEL = r"C:\ELITEC\VALIDACION_APLICACION"
    else:
        CARPETA_EXCEL = os.path.join(app.root_path, "uploads")
EXTENSIONES_PERMITIDAS = {".xlsx", ".xlsm", ".xls"}
COLUMNAS_REQUERIDAS = (
    "codigo",
    "producto",
    "fecha",
    "cantidad",
    "precio",
    "subtotal",
)
ALIAS_COLUMNAS = {
    "codigo": {
        "codigo",
        "codigoproducto",
        "codigodeproducto",
        "codproducto",
        "idproducto",
        "id",
        "cod",
    },
    "producto": {
        "producto",
        "nombre",
        "nombreproducto",
        "nombredelproducto",
        "descripcion",
        "detalle",
        "item",
    },
    "fecha": {"fecha", "fechaventa", "fechadeventa", "dia"},
    "cantidad": {"cantidad", "cantidadvendida", "unidades", "cant"},
    "precio": {"precio", "preciounitario", "valorunitario", "pu"},
    "subtotal": {"subtotal", "importe", "importetotal", "total", "monto"},
}
MESES = (
    "enero",
    "febrero",
    "marzo",
    "abril",
    "mayo",
    "junio",
    "julio",
    "agosto",
    "septiembre",
    "octubre",
    "noviembre",
    "diciembre",
)

lista_productos = []


class ErrorImportacion(ValueError):
    def __init__(self, mensajes):
        if isinstance(mensajes, str):
            mensajes = [mensajes]
        self.mensajes = mensajes
        super().__init__(" ".join(mensajes))


def normalizar_columna(nombre):
    texto = unicodedata.normalize("NFKD", str(nombre))
    texto = "".join(
        caracter for caracter in texto if not unicodedata.combining(caracter)
    )
    texto = re.sub(r"(?i)s\s*/", "", texto)
    return re.sub(r"[^a-z0-9]+", "", texto.lower())


def mapear_columnas(columnas):
    mapa = {}
    for columna in columnas:
        normalizada = normalizar_columna(columna)
        for canonica, alias in ALIAS_COLUMNAS.items():
            if normalizada in alias and canonica not in mapa:
                mapa[canonica] = columna
                break
    faltantes = [
        columna for columna in COLUMNAS_REQUERIDAS if columna not in mapa
    ]
    return mapa, faltantes


def convertir_codigo(valor):
    if pd.isna(valor):
        raise ValueError("el código está vacío")
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        numero = float(valor)
        if numero.is_integer():
            return int(numero)
        return numero
    texto = str(valor).strip()
    if not texto:
        raise ValueError("el código está vacío")
    if texto.isdigit() and not texto.startswith("0"):
        return int(texto)
    return texto


def convertir_producto(valor):
    if pd.isna(valor):
        raise ValueError("el nombre del producto está vacío")
    texto = str(valor).strip()
    if not texto:
        raise ValueError("el nombre del producto está vacío")
    return texto


def convertir_fecha(valor):
    if pd.isna(valor):
        raise ValueError("la fecha está vacía")
    if isinstance(valor, (datetime, date)):
        return pd.Timestamp(valor).date()
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        numero = int(valor)
        if 19000101 <= numero <= 21001231:
            return datetime.strptime(str(numero), "%Y%m%d").date()
        if 1 <= numero <= 2958465:
            return (datetime(1899, 12, 30) + timedelta(days=numero)).date()
        raise ValueError("fecha inválida")
    texto = str(valor).strip()
    fecha = pd.to_datetime(texto, dayfirst=True, errors="coerce")
    if pd.isna(fecha):
        raise ValueError(
            f"fecha inválida ({texto!r}); use el formato DD/MM/AAAA"
        )
    return fecha.date()


def convertir_numero(valor, nombre_campo, permitir_cero=True):
    if pd.isna(valor):
        raise ValueError(f"{nombre_campo} está vacío")
    if isinstance(valor, (int, float)) and not isinstance(valor, bool):
        numero = float(valor)
    else:
        texto = str(valor).strip()
        texto = re.sub(r"(?i)s\s*/", "", texto)
        texto = texto.replace(",", "")
        texto = re.sub(r"[^\d.\-]", "", texto)
        numero = pd.to_numeric(texto, errors="coerce")
        if pd.isna(numero):
            raise ValueError(
                f"{nombre_campo} debe ser un número "
                f"(valor recibido: {valor!r})"
            )
        numero = float(numero)
    if not math.isfinite(numero):
        raise ValueError(f"{nombre_campo} debe ser un número válido")
    if numero < 0:
        raise ValueError(f"{nombre_campo} no puede ser negativo")
    if not permitir_cero and numero == 0:
        raise ValueError(f"{nombre_campo} debe ser mayor que cero")
    return numero


def convertir_cantidad(valor):
    numero = convertir_numero(valor, "la cantidad", permitir_cero=False)
    return int(numero) if numero.is_integer() else numero


def convertir_precio(valor, nombre_campo):
    return round(convertir_numero(valor, nombre_campo), 2)


def leer_hoja_productos(ruta):
    extension = os.path.splitext(ruta)[1].lower()
    if extension not in EXTENSIONES_PERMITIDAS:
        raise ErrorImportacion(
            "Formato no permitido. Use un archivo .xlsx, .xlsm o .xls."
        )
    if extension == ".xls":
        try:
            import xlrd  # noqa: F401
        except ImportError as error:
            raise ErrorImportacion(
                "Para leer archivos .xls instale la librería xlrd."
            ) from error
        engine = "xlrd"
    else:
        engine = "openpyxl"

    try:
        with pd.ExcelFile(ruta, engine=engine) as libro:
            hoja = next(
                (
                    nombre
                    for nombre in libro.sheet_names
                    if normalizar_columna(nombre) == "productos"
                ),
                None,
            )
            if hoja is None:
                hojas = (
                    ", ".join(str(nombre) for nombre in libro.sheet_names)
                    or "ninguna"
                )
                raise ErrorImportacion(
                    'El archivo debe contener una hoja llamada "Productos". '
                    f"Hojas encontradas: {hojas}."
                )
            datos = libro.parse(hoja, dtype=object)
    except ErrorImportacion:
        raise
    except Exception as error:
        raise ErrorImportacion(
            "No se pudo leer el archivo. Verifique que sea un Excel válido "
            "y no esté dañado."
        ) from error
    return datos


def construir_productos(datos):
    mapa, faltantes = mapear_columnas(datos.columns)
    if faltantes:
        encontradas = (
            ", ".join(str(columna) for columna in datos.columns) or "ninguna"
        )
        raise ErrorImportacion(
            "Faltan columnas obligatorias: "
            + ", ".join(faltantes)
            + ". Columnas encontradas: "
            + encontradas
            + "."
        )

    productos = []
    errores = []
    for posicion, (_, fila) in enumerate(datos.iterrows(), start=2):
        if fila.isna().all():
            continue
        try:
            productos.append(
                Producto(
                    idProducto=convertir_codigo(fila[mapa["codigo"]]),
                    producto=convertir_producto(fila[mapa["producto"]]),
                    fecha=convertir_fecha(fila[mapa["fecha"]]),
                    cantidad=convertir_cantidad(fila[mapa["cantidad"]]),
                    precio=convertir_precio(
                        fila[mapa["precio"]], "el precio"
                    ),
                    subtotal=convertir_precio(
                        fila[mapa["subtotal"]], "el subtotal"
                    ),
                )
            )
        except (ValueError, TypeError, OverflowError) as error:
            errores.append(f"Fila {posicion}: {error}.")

    if errores:
        resumen = errores[:10]
        if len(errores) > 10:
            resumen.append(f"... y {len(errores) - 10} error(es) más.")
        raise ErrorImportacion(resumen)
    if not productos:
        raise ErrorImportacion(
            "La hoja Productos no contiene filas de datos."
        )
    return productos


def obtener_carpeta_excel():
    try:
        os.makedirs(CARPETA_EXCEL, exist_ok=True)
        return CARPETA_EXCEL
    except OSError:
        carpeta_local = os.path.join(app.root_path, "uploads")
        os.makedirs(carpeta_local, exist_ok=True)
        return carpeta_local


def obtener_periodo():
    if not lista_productos:
        return ""
    fechas = [producto.fecha for producto in lista_productos]
    inicio = min(fechas)
    fin = max(fechas)
    if inicio.year == fin.year and inicio.month == fin.month:
        return f"{MESES[inicio.month - 1].capitalize()} {inicio.year}"
    return f"{inicio.strftime('%d/%m/%Y')} al {fin.strftime('%d/%m/%Y')}"


def clave_codigo(valor):
    texto = str(valor)
    if texto.isdigit():
        return (0, int(texto))
    return (1, texto)


def crear_plantilla_excel():
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.worksheet.datavalidation import DataValidation
    from openpyxl.worksheet.table import Table, TableStyleInfo

    libro = Workbook()
    hoja = libro.active
    hoja.title = "Productos"

    hoja.append(list(COLUMNAS_REQUERIDAS))
    hoja.append([1, "Producto de ejemplo", date(2026, 8, 15), 2, 10.5, 21.0])

    for celda in hoja[1]:
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="FF4F46E5")
        celda.alignment = Alignment(horizontal="center", vertical="center")

    hoja["C2"].number_format = "DD/MM/YYYY"
    hoja["E2"].number_format = "#,##0.00"
    hoja["F2"].number_format = "#,##0.00"

    tabla = Table(displayName="TablaProductos", ref="A1:F2")
    tabla.tableStyleInfo = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False,
    )
    hoja.add_table(tabla)
    hoja.freeze_panes = "A2"
    hoja.column_dimensions["A"].width = 12
    hoja.column_dimensions["B"].width = 28
    hoja.column_dimensions["C"].width = 14
    hoja.column_dimensions["D"].width = 12
    hoja.column_dimensions["E"].width = 14
    hoja.column_dimensions["F"].width = 14

    validacion_cantidad = DataValidation(
        type="decimal",
        operator="greaterThan",
        formula1="0",
        allow_blank=True,
    )
    validacion_cantidad.errorTitle = "Cantidad inválida"
    validacion_cantidad.error = "Ingrese una cantidad mayor que cero."
    hoja.add_data_validation(validacion_cantidad)
    validacion_cantidad.add("D2:D1000")

    validacion_monto = DataValidation(
        type="decimal",
        operator="greaterThanOrEqual",
        formula1="0",
        allow_blank=True,
    )
    validacion_monto.errorTitle = "Monto inválido"
    validacion_monto.error = "Ingrese un número mayor o igual a cero."
    hoja.add_data_validation(validacion_monto)
    validacion_monto.add("E2:F1000")

    instrucciones = libro.create_sheet("Instrucciones")
    instrucciones["A1"] = "Formato para importar ventas"
    instrucciones["A1"].font = Font(bold=True, size=14, color="4F46E5")

    lineas = [
        "",
        "1. Use la hoja 'Productos'.",
        "2. Mantenga los encabezados de la fila 1: codigo, producto, fecha, "
        "cantidad, precio, subtotal.",
        "3. Cada fila debe representar una venta.",
        "4. La fecha debe tener formato DD/MM/AAAA (ejemplo: 15/08/2026).",
        "5. cantidad debe ser mayor que cero; precio y subtotal deben ser "
        "mayores o iguales a cero.",
        "6. No use celdas combinadas ni filas de totales.",
        "7. Reemplace o elimine la fila de ejemplo antes de importar.",
        "8. Guarde el archivo en formato .xlsx, .xlsm o .xls.",
        "",
        "Ejemplo:",
        "codigo | producto | fecha | cantidad | precio | subtotal",
        "1 | Producto de ejemplo | 15/08/2026 | 2 | 10.50 | 21.00",
    ]
    for indice, linea in enumerate(lineas, start=2):
        instrucciones.cell(row=indice, column=1, value=linea)
    instrucciones.column_dimensions["A"].width = 110
    for fila in instrucciones.iter_rows(
        min_row=2, max_row=len(lineas) + 1, min_col=1, max_col=1
    ):
        for celda in fila:
            celda.alignment = Alignment(wrap_text=True, vertical="top")

    buffer = io.BytesIO()
    libro.save(buffer)
    buffer.seek(0)
    return buffer


@app.errorhandler(413)
def archivo_muy_grande(error):
    flash("El archivo supera el límite de 20 MB.", "error")
    return redirect(url_for("listaproductos"))


@app.route("/")
def inicio():
    return render_template("inicio.html", periodo=obtener_periodo())


@app.route("/importar_excel", methods=["POST"])
def leer_excel():
    archivo = request.files.get("archivo_a_subir")
    if archivo is None or archivo.filename == "":
        flash("Seleccione un archivo Excel antes de importar.", "error")
        return redirect(url_for("listaproductos"))

    nombre_archivo = secure_filename(archivo.filename) or "ventas.xlsx"
    extension = os.path.splitext(nombre_archivo)[1].lower()
    if extension not in EXTENSIONES_PERMITIDAS:
        flash(
            "Formato no permitido. Use un archivo .xlsx, .xlsm o .xls.",
            "error",
        )
        return redirect(url_for("listaproductos"))

    try:
        carpeta = obtener_carpeta_excel()
        ruta_final = os.path.join(carpeta, nombre_archivo)
        archivo.save(ruta_final)
    except OSError:
        flash(
            "No se pudo guardar el archivo. Verifique los permisos de la "
            "carpeta de carga.",
            "error",
        )
        return redirect(url_for("listaproductos"))

    try:
        datos = leer_hoja_productos(ruta_final)
        productos = construir_productos(datos)
    except ErrorImportacion as error:
        for mensaje in error.mensajes:
            flash(mensaje, "error")
        return redirect(url_for("listaproductos"))

    lista_productos.clear()
    lista_productos.extend(productos)
    flash(
        f"Se importaron {len(productos)} filas correctamente.",
        "exito",
    )
    return redirect(url_for("listaproductos"))


@app.route("/formato")
def descargar_formato():
    return send_file(
        crear_plantilla_excel(),
        as_attachment=True,
        download_name="formato_ventas.xlsx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument."
            "spreadsheetml.sheet"
        ),
    )


@app.route("/listaproductos")
def listaproductos():
    ventas = obtener_resumen_ventas()
    total_cantidad = sum(venta["cantidad"] for venta in ventas)
    total_monto = sum(venta["subtotal"] for venta in ventas)
    return render_template(
        "listaproductos.html",
        productos=ventas,
        total_cantidad=total_cantidad,
        total_monto=total_monto,
        periodo=obtener_periodo(),
    )


def obtener_resumen_ventas():
    resumen = {}
    for producto in lista_productos:
        clave = (producto.fecha, producto.producto)
        if clave in resumen:
            resumen[clave]["cantidad"] += producto.cantidad
            resumen[clave]["subtotal"] += producto.subtotal
        else:
            resumen[clave] = {
                "codigo": producto.idProducto,
                "producto": producto.producto,
                "fecha": producto.fecha,
                "cantidad": producto.cantidad,
                "precio": producto.precio,
                "subtotal": producto.subtotal,
            }
    return [resumen[clave] for clave in sorted(resumen)]


@app.route("/resumen")
def resumen():
    ventas = obtener_resumen_ventas()

    por_fecha = {}
    for venta in ventas:
        por_fecha.setdefault(venta["fecha"], []).append(venta)

    fecha_grupos = []
    for fecha in sorted(por_fecha):
        productos = sorted(
            por_fecha[fecha], key=lambda venta: clave_codigo(venta["codigo"])
        )
        fecha_grupos.append({"fecha": fecha, "productos": productos})

    total_cantidad = sum(venta["cantidad"] for venta in ventas)
    total_monto = sum(venta["subtotal"] for venta in ventas)

    return render_template(
        "resumen.html",
        grupos=fecha_grupos,
        total_cantidad=total_cantidad,
        total_monto=total_monto,
        periodo=obtener_periodo(),
    )


@app.route("/estadistica")
def estadistica():
    periodo = obtener_periodo()
    if not lista_productos:
        return render_template(
            "estadistica.html",
            grafico="",
            mensaje="Primero importe el archivo Excel de ventas",
            periodo=periodo,
        )

    totales = {}
    for producto in lista_productos:
        if producto.producto in totales:
            totales[producto.producto]["cantidad"] += producto.cantidad
            totales[producto.producto]["subtotal"] += producto.subtotal
        else:
            totales[producto.producto] = {
                "codigo": producto.idProducto,
                "cantidad": producto.cantidad,
                "subtotal": producto.subtotal,
            }

    top5 = sorted(
        totales.items(),
        key=lambda elemento: (-elemento[1]["cantidad"], elemento[0]),
    )[:5]

    nombres = [nombre for nombre, _ in top5][::-1]
    cantidades = [datos["cantidad"] for _, datos in top5][::-1]
    codigos = [str(datos["codigo"]) for _, datos in top5][::-1]
    subtotales = [datos["subtotal"] for _, datos in top5][::-1]

    barras = px.bar(
        x=cantidades,
        y=nombres,
        orientation="h",
        text=cantidades,
        custom_data=[codigos, subtotales],
        color=cantidades,
        color_continuous_scale=["#c7d2fe", "#4f46e5"],
    )
    barras.update_traces(
        texttemplate="%{text:,.0f}",
        textposition="outside",
        hovertemplate=(
            "<b>%{y}</b><br>"
            "Código: %{customdata[0]}<br>"
            "Cantidad vendida: %{x:,.0f} unidades<br>"
            "Importe total: S/ %{customdata[1]:,.2f}"
            "<extra></extra>"
        ),
        marker_line_width=0,
        cliponaxis=False,
    )
    barras.update_layout(
        template="plotly_white",
        height=420,
        margin=dict(l=10, r=70, t=20, b=20),
        font=dict(
            family="'Segoe UI', system-ui, sans-serif",
            size=14,
            color="#1e293b",
        ),
        xaxis=dict(
            title="Cantidad vendida (unidades)",
            showgrid=True,
            gridcolor="#e2e8f0",
            zeroline=False,
            rangemode="tozero",
            tickformat=",.0f",
        ),
        yaxis=dict(title="", showgrid=False, automargin=True),
        showlegend=False,
        coloraxis_showscale=False,
        bargap=0.35,
        barcornerradius=6,
        hoverlabel=dict(
            bgcolor="#ffffff",
            bordercolor="#e2e8f0",
            font=dict(family="'Segoe UI', system-ui, sans-serif", size=13),
        ),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    grafico = pio.to_html(
        barras,
        full_html=False,
        config={"displayModeBar": False, "responsive": True},
    )

    return render_template(
        "estadistica.html",
        grafico=grafico,
        periodo=periodo,
    )


if __name__ == "__main__":
    app.run(debug=True)
