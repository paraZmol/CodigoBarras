from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import os

def dibujar_titulo_prueba():
    nombre_archivo = "prueba_titulo_biblioteca.pdf"
    
    # fuente open sans
    ruta_fuente = "OpenSans-Bold.ttf"
    
    # verificacion de existencia de fuente
    if os.path.exists(ruta_fuente):
        pdfmetrics.registerFont(TTFont('OpenSans-Bold', ruta_fuente))
        fuente_usada = "OpenSans-Bold"
    else:
        # en caso de no estar usar Helvetica para que no falle
        print(f"AVISO: No se encontró '{ruta_fuente}'. Asegúrate de tener el archivo en la carpeta.")
        fuente_usada = "Helvetica-Bold"

    c = canvas.Canvas(nombre_archivo, pagesize=A4)
    
    # dimensiones de la hoja A4
    ancho_hoja, alto_hoja = A4
    
    # medidas
    posicion_x_titulo = 5.52 * cm
    ancho_titulo = 9.97 * cm
    alto_titulo = 1.18 * cm
    
    # margen desde arriba
    margen_superior_y = 1.1 * cm 
    
    # calcular la posision desde abajo
    posicion_y_titulo = alto_hoja - margen_superior_y - alto_titulo

    # cuadro rojo de referencia
    c.setStrokeColorRGB(1, 0, 0)
    c.setLineWidth(1)
    c.rect(posicion_x_titulo, posicion_y_titulo, ancho_titulo, alto_titulo)

    # texto negro
    c.setFillColorRGB(0, 0, 0)
    
    # centro horizontal
    centro_x_titulo = posicion_x_titulo + (ancho_titulo / 2)
    
    # configuracion de la feunte
    tamano_fuente = 13 
    c.setFont(fuente_usada, tamano_fuente)

    # linea 1
    altura_linea_1 = posicion_y_titulo + alto_titulo - 0.35*cm
    c.drawCentredString(centro_x_titulo, altura_linea_1, "BIBLIOTECA FEC - UBICACIÓN ESTANTERIA")
    
    # linea 2
    altura_linea_2 = posicion_y_titulo + 0.25*cm
    c.drawCentredString(centro_x_titulo, altura_linea_2, "339.1 - 343.08")

    c.showPage()
    c.save()
    print(f"Generado correctamente: {nombre_archivo}")

if __name__ == "__main__":
    dibujar_titulo_prueba()