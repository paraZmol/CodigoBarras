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

# dibujar cuadro segun la posicion x y
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
    nombre_archivo = "etiqueta_completa5.pdf"
    
    c = canvas.Canvas(nombre_archivo, pagesize=A4)
    ancho_hoja, alto_hoja = A4
    
    # configuracio nde fuente
    fuente_actual = obtener_fuente()
    
    # deibujar el titulo fijo
    dibujar_titulo(c, ancho_hoja, alto_hoja, fuente_actual)
    
    # definimos alto del cuadro para calcular la y de reportlab
    alto_cuadro = 3.28 * cm
    
    # definimos un espacio visual entre filas (1 milimetro)
    espacio_entre_filas = 0.05 * cm
    
    # definimos la Y inicial de la primera fila
    y_inicial = 2.73 * cm

    # calculo de la segunda fila
    fila_2_y = y_inicial + alto_cuadro + espacio_entre_filas
    
    # inclusion de espacio
    lista_filas_y = [y_inicial, fila_2_y]
    
    # lista de ubicaciones horizontales (x)
    ubicaciones_x = [0.3 * cm, 7.05 * cm, 13.79 * cm]
    
    # bucle que recorre las filas
    for y_arriba in lista_filas_y:
        
        # calculamos la y real para esta fila completa
        posicion_y_real = alto_hoja - y_arriba - alto_cuadro
        
        # bucle que recorre las columnas de esta fila
        for x in ubicaciones_x:
            dibujar_cuadro(c, x, posicion_y_real)
    
    # guardar el pdf
    c.showPage()
    c.save()
    print(f"Generado correctamente: {nombre_archivo}")

if __name__ == "__main__":
    generar_etiqueta_completa()