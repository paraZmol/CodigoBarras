from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm

def dibujar_titulo_prueba():
    nombre_archivo = "prueba_titulo.pdf"
    c = canvas.Canvas(nombre_archivo, pagesize=A4)
    
    # dimensiones de hoja
    ancho_hoja, alto_hoja = A4
    
    # medidas de ltitulo
    posicion_x_titulo = 5.52 * cm
    ancho_titulo = 9.97 * cm
    alto_titulo = 1.18 * cm
    
    # coordenadas
    margen_superior_y = 1.1 * cm 
    posicion_y_titulo = alto_hoja - margen_superior_y - alto_titulo

    # borde de cuadro rojo
    c.setStrokeColorRGB(1, 0, 0)
    c.setLineWidth(1)
    c.rect(posicion_x_titulo, posicion_y_titulo, ancho_titulo, alto_titulo)

    # texto dentro del cuador
    c.setFillColorRGB(0, 0, 0)
    centro_x_titulo = posicion_x_titulo + (ancho_titulo / 2)
    
    # fuente de texto
    fuente_usada = "OpenSans-Bold" 
    tamano_fuente = 13
    c.setFont(fuente_usada, tamano_fuente) 

    # dibujo de texto
    c.drawCentredString(centro_x_titulo, posicion_y_titulo + 0.3*cm, "MODERNIZACIÓN DE LA INFRAESTRUCTURA")
    
    c.showPage()
    c.save()
    print(f"Generado: {nombre_archivo}")

if __name__ == "__main__":
    dibujar_titulo_prueba()