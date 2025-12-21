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
def dibujar_titulo(c, ancho_hoja, alto_hoja, fuente_usada):
    """Dibuja solo la parte del título en el canvas 'c'."""
    
    # medidas de titulo
    posicion_x_titulo = 5.52 * cm
    ancho_titulo = 9.97 * cm
    alto_titulo = 1.18 * cm
    margen_superior_y = 1.1 * cm 
    
    # poscion de Y 
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

# metodo 2, dibujar el cuerpo -- solo cuadro
def dibujar_cuerpo(c, ancho_hoja, alto_hoja):
    """Dibuja los cuadros del cuerpo en el canvas 'c'."""
    
    # medidas del cuerpo para todos
    ancho_cuadros = 6.56 * cm
    
    # definimos alto igual para todos segun instrucciones
    alto_cuadros = 3.28 * cm
    
    # posicion y desde arriba segun indicacion 2.73
    y_arriba = 2.73 * cm

    # calcular la posicion y desde abajo para reportlab
    posicion_y = alto_hoja - y_arriba - alto_cuadros

    c.setLineWidth(1)
    c.setStrokeColorRGB(0, 0, 0)

    # cuadro 1
    # ubicacion x 0.3 y 2.73
    c.rect(0.3 * cm, posicion_y, ancho_cuadros, alto_cuadros)
    
    # cuadro 2
    # ubicacion x 7.05 y 2.73
    c.rect(7.05 * cm, posicion_y, ancho_cuadros, alto_cuadros)
    
    # cuadro 3
    # ubicacion x 13.79 y 2.73
    c.rect(13.79 * cm, posicion_y, ancho_cuadros, alto_cuadros)

# funcion principal
def generar_etiqueta_completa():
    nombre_archivo = "etiqueta_completa2.pdf"
    
    c = canvas.Canvas(nombre_archivo, pagesize=A4)
    ancho_hoja, alto_hoja = A4
    
    # configuracio nde fuente
    fuente_actual = obtener_fuente()
    
    # deibujar el titulo y el cuerpo
    dibujar_titulo(c, ancho_hoja, alto_hoja, fuente_actual)
    dibujar_cuerpo(c, ancho_hoja, alto_hoja)
    
    # guardar el pdf
    c.showPage()
    c.save()
    print(f"Generado correctamente: {nombre_archivo}")

if __name__ == "__main__":
    generar_etiqueta_completa()