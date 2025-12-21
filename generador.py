from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

# funcion de la fuente
def obtener_fuente():
    """Registra la fuente Open Sans si existe, sino devuelve Helvetica."""
    ruta_fuente = "OpenSans-Bold.ttf"
    if os.path.exists(ruta_fuente):
        pdfmetrics.registerFont(TTFont('OpenSans-Bold', ruta_fuente))
        return "OpenSans-Bold"
    else:
        print(f"AVISO: No se encontró '{ruta_fuente}'. Usando Helvetica.")
        return "Helvetica-Bold"

# dibujar titulo
def dibujar_titulo(c, ancho_hoja, alto_hoja, fuente_usada, margen_superior_y):
    """Dibuja solo la parte del título en el canvas 'c'."""
    
    # medidas de titulo 
    posicion_x_titulo = 5.52 * cm
    ancho_titulo = 9.97 * cm
    alto_titulo = 1.18 * cm
    
    # poscion de Y calculada con el margen variable
    posicion_y_titulo = alto_hoja - margen_superior_y - alto_titulo

    # cuadro rojo
    #c.setStrokeColorRGB(1, 0, 0) 
    #c.setLineWidth(1)
    #c.rect(posicion_x_titulo, posicion_y_titulo, ancho_titulo, alto_titulo)

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

# dibujo de cuadros segun x y
def dibujar_cuadro(c, x, y):
    """Recibe la ubicación (x, y) y dibuja un cuadro."""
    
    # medidas constantes del cuadro
    ancho_cuadro = 6.56 * cm
    alto_cuadro = 3.28 * cm

    c.setLineWidth(1)
    c.setStrokeColorRGB(0, 0, 0)

    # dibujar el rectangulo dinamico
    c.rect(x, y, ancho_cuadro, alto_cuadro)

# funcion principal
def generar_etiqueta_completa():
    nombre_archivo = "etiqueta_completa8.pdf"
    
    c = canvas.Canvas(nombre_archivo, pagesize=A4)
    ancho_hoja, alto_hoja = A4
    fuente_actual = obtener_fuente()
    
    # ***************************** configuracion de margenes de impresion
    
    margen_superior_papel = 1.1 * cm      # margen superior -- tituño
    margen_izquierdo_papel = 0.4 * cm     # margen isquierdo -- primera columna
    
    espacio_entre_columnas = 0.15 * cm    # espaciado horizontal
    espacio_entre_filas = 0.06 * cm       # espaciado vertical
    
    y_inicial_grid = 2.70 * cm            # inicio del primer cuadro
    
    # dibujo de titulo
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
    for y_arriba in lista_filas_y:
        # calcular la nueva coordenada Y segun ReportLab
        posicion_y_real = alto_hoja - y_arriba - alto_cuadro
        
        for x in ubicaciones_x:
            dibujar_cuadro(c, x, posicion_y_real)
    
    c.showPage()
    c.save()
    print(f"Generado correctamente: {nombre_archivo}")

if __name__ == "__main__":
    generar_etiqueta_completa()