from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

# libreria code 39
from reportlab.graphics.barcode import code39
from reportlab.graphics.shapes import Drawing

# funcion de la fuente
def obtener_fuente():
    ruta_fuente = "OpenSans-Bold.ttf"
    if os.path.exists(ruta_fuente):
        pdfmetrics.registerFont(TTFont('OpenSans-Bold', ruta_fuente))
        return "OpenSans-Bold"
    else:
        print(f"AVISO: No se encontró '{ruta_fuente}'. Usando Helvetica.")
        return "Helvetica-Bold"

# dibujar titulo principal de la hoja
def dibujar_titulo(c, ancho_hoja, alto_hoja, fuente_usada, margen_superior_y):
    """Dibuja solo la parte del título principal en el canvas 'c'."""
    
    # medidas de titulo 
    posicion_x_titulo = 5.52 * cm
    ancho_titulo = 9.97 * cm
    alto_titulo = 1.18 * cm
    
    # poscion de Y calculada con el margen variable
    posicion_y_titulo = alto_hoja - margen_superior_y - alto_titulo

    # text negro
    c.setFillColorRGB(0, 0, 0) 
    centro_x_titulo = posicion_x_titulo + (ancho_titulo / 2)
    
    tamano_fuente = 13 
    c.setFont(fuente_usada, tamano_fuente)

    # linea 1 de titulo
    altura_linea_1 = posicion_y_titulo + alto_titulo - 0.35*cm
    c.drawCentredString(centro_x_titulo, altura_linea_1, "BIBLIOTECA FEC - UBICACIÓN ESTANTERIA")
    
    # linea 2 de titulo
    altura_linea_2 = posicion_y_titulo + 0.25*cm
    c.drawCentredString(centro_x_titulo, altura_linea_2, "339.1 - 343.08")

# dibujo de cuadros
def dibujar_cuadro(c, x, y, fuente, ajuste_y_barras, valor_codigo="12345"):
    
    # medidas constantes del cuadro
    ancho_cuadro = 6.56 * cm
    alto_cuadro = 3.28 * cm

    # borde de cuadro
    c.setLineWidth(1)
    c.setStrokeColorRGB(0, 0, 0)
    c.setFillColorRGB(0, 0, 0) 
    c.rect(x, y, ancho_cuadro, alto_cuadro)

    # titulo interno
    c.setFont(fuente, 10) 
    titulo_interno = "Sistema de Bibliotecas UNASAM"
    
    # calculo y para el titulo de cada cuadro
    alto_titulo_cuadro = 0.4*cm 
    y_texto = y + alto_cuadro - alto_titulo_cuadro - 0.2*cm
    
    centro_x_cuadro = x + (ancho_cuadro / 2)
    c.drawCentredString(centro_x_cuadro, y_texto, titulo_interno)

    # **************** configuracion visual y textual
    
    # altura total entre ambos
    altura_total_visual = 1.34 * cm
    
    # medida de cada parte para un amximo de 1.34
    alto_barras = 0.94 * cm
    espacio_texto = altura_total_visual - alto_barras # aprox 0.4 cm
    
    # calcular la base Y del bloque
    # formula: Y_base_cuadro + (Espacio libre / 2) + Ajuste manual
    y_base_bloque = y + (alto_cuadro - altura_total_visual) / 2 + ajuste_y_barras

    # dibujar el texto
    c.setFont(fuente, 9)
    
    # formato de texto
    texto_a_mostrar = f"*{valor_codigo}*" 
    
    # ligeramente por encima de la base
    y_texto_codigo = y_base_bloque + 0.1 * cm 
    c.drawCentredString(centro_x_cuadro, y_texto_codigo, texto_a_mostrar)

    # dibujar las barras
    ancho_barras = 0.045 * cm 
    
    barcode = code39.Standard39(
        valor_codigo, 
        barHeight=alto_barras, 
        barWidth=ancho_barras, 
        checksum=0,
        humanReadable=False  # Desactivamos el automático para usar el nuestro manual
    )
    
    ancho_codigo_real = barcode.width
    x_barcode = x + (ancho_cuadro / 2) - (ancho_codigo_real / 2)
    
    # la posicion Y del codigo de barras es justo encima del espacio reservado para texto
    y_barcode = y_base_bloque + espacio_texto + 0.03*cm

    barcode.drawOn(c, x_barcode, y_barcode)


# funcion principal
def generar_etiqueta_completa():
    nombre_archivo = "codigo_visual_numero2.pdf"
    
    c = canvas.Canvas(nombre_archivo, pagesize=A4)
    ancho_hoja, alto_hoja = A4
    fuente_actual = obtener_fuente()
    
    # ***************************** configuracion de margenes y ajustes
    
    margen_superior_papel = 1.1 * cm      # margen superior
    margen_izquierdo_papel = 0.4 * cm     # margen isquierdo
    
    espacio_entre_columnas = 0.15 * cm    # espaciado horizontal
    espacio_entre_filas = 0.06 * cm       # espaciado vertical
    
    y_inicial_grid = 2.70 * cm            # inicio del primer cuadro
    
    # variable para mover el bloque completo (barras + numero)
    ajuste_vertical_codigo = 0.1 * cm 
    
    # dibujo de titulo principal de hoja
    dibujar_titulo(c, ancho_hoja, alto_hoja, fuente_actual, margen_superior_papel)
    
    #*********************** logica de cuadros
    
    alto_cuadro = 3.28 * cm
    ancho_cuadro = 6.56 * cm
    numero_de_filas = 8
    numero_de_columnas = 3

    # calculo de posiciones Y
    lista_filas_y = []
    posicion_y_actual = y_inicial_grid
    
    for i in range(numero_de_filas):
        lista_filas_y.append(posicion_y_actual)
        posicion_y_actual = posicion_y_actual + alto_cuadro + espacio_entre_filas

    # calculo posiciones X
    ubicaciones_x = []
    posicion_x_actual = margen_izquierdo_papel
    
    for i in range(numero_de_columnas):
        ubicaciones_x.append(posicion_x_actual)
        posicion_x_actual = posicion_x_actual + ancho_cuadro + espacio_entre_columnas

    # dibujar grilla
    contador = 1
    for y_arriba in lista_filas_y:
        posicion_y_real = alto_hoja - y_arriba - alto_cuadro
        
        for x in ubicaciones_x:
            codigo_prueba = f"1000{contador}" 
            dibujar_cuadro(c, x, posicion_y_real, fuente_actual, ajuste_vertical_codigo, valor_codigo=codigo_prueba)
            contador += 1
    
    c.showPage()
    c.save()
    print(f"Generado correctamente: {nombre_archivo}")

if __name__ == "__main__":
    generar_etiqueta_completa()