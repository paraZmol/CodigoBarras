from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.barcode import code39
import os

# ******************* ZONA DE CONFIGURACION *******************
class Config:
    # archivo de salida
    NOMBRE_ARCHIVO = "codigo_visual_numero3.pdf"
    
    # fuente personalizada
    RUTA_FUENTE = "OpenSans-Bold.ttf"
    
    # margenes principales
    MARGEN_SUPERIOR = 1.1 * cm
    MARGEN_IZQUIERDO = 0.4 * cm
    
    # espaciado entre cuadros
    ESPACIO_HORIZONTAL = 0.15 * cm  # espacio entre columnas
    ESPACIO_VERTICAL = 0.06 * cm    # espacio entre filas
    
    # posicion inicial del grid
    Y_INICIAL_GRID = 2.70 * cm
    
    # ajuste fino del codigo de barras dentro del cuadro
    AJUSTE_VERTICAL_CODIGO = 0.1 * cm
    
    # dimensiones del cuadro individual
    ANCHO_CUADRO = 6.56 * cm
    ALTO_CUADRO = 3.28 * cm
    
    # cantidad de cuadros
    FILAS = 8
    COLUMNAS = 3
    
    # configuracion del titulo principal
    TITULO_LINEA1 = "BIBLIOTECA FEC - UBICACION ESTANTERIA"
    TITULO_LINEA2 = "339.1 - 343.08"
    TAMANO_FUENTE_TITULO = 13
    
    # configuracion del cuadro individual
    TITULO_CUADRO = "Sistema de Bibliotecas UNASAM"
    TAMANO_FUENTE_CUADRO = 10
    TAMANO_FUENTE_CODIGO = 9
    
    # configuracion del codigo de barras
    ANCHO_BARRAS = 0.045 * cm
    ALTO_BARRAS = 0.94 * cm
    
    # codigo inicial (se incrementara automaticamente)
    CODIGO_INICIAL = 10001

# *************************************************************


class GeneradorEtiquetas:
    def __init__(self):
        self.config = Config()
        self.fuente = self._cargar_fuente()
        
    def _cargar_fuente(self):
        """carga la fuente personalizada o usa la default"""
        if os.path.exists(self.config.RUTA_FUENTE):
            pdfmetrics.registerFont(TTFont('OpenSans-Bold', self.config.RUTA_FUENTE))
            return "OpenSans-Bold"
        else:
            print(f"aviso: no se encontro '{self.config.RUTA_FUENTE}', usando helvetica")
            return "Helvetica-Bold"
    
    def _dibujar_titulo_principal(self, c, ancho_hoja, alto_hoja):
        """dibuja el titulo principal en la parte superior"""
        # dimensiones del titulo
        pos_x = 5.52 * cm
        ancho = 9.97 * cm
        alto = 1.18 * cm
        
        # calcular posicion y
        pos_y = alto_hoja - self.config.MARGEN_SUPERIOR - alto
        centro_x = pos_x + (ancho / 2)
        
        # configurar texto
        c.setFillColorRGB(0, 0, 0)
        c.setFont(self.fuente, self.config.TAMANO_FUENTE_TITULO)
        
        # linea 1
        y_linea1 = pos_y + alto - 0.35 * cm
        c.drawCentredString(centro_x, y_linea1, self.config.TITULO_LINEA1)
        
        # linea 2
        y_linea2 = pos_y + 0.25 * cm
        c.drawCentredString(centro_x, y_linea2, self.config.TITULO_LINEA2)
    
    def _dibujar_codigo_barras(self, c, x_centro, y_base, codigo):
        """dibuja el codigo de barras en la posicion especificada"""
        barcode = code39.Standard39(
            codigo,
            barHeight=self.config.ALTO_BARRAS,
            barWidth=self.config.ANCHO_BARRAS,
            checksum=0,
            humanReadable=False
        )
        
        # centrar el codigo de barras
        ancho_codigo = barcode.width
        x_barcode = x_centro - (ancho_codigo / 2)
        
        barcode.drawOn(c, x_barcode, y_base)
    
    def _dibujar_texto_codigo(self, c, x_centro, y_base, codigo):
        """dibuja el texto del codigo debajo del codigo de barras"""
        c.setFont(self.fuente, self.config.TAMANO_FUENTE_CODIGO)
        texto = f"*{codigo}*"
        c.drawCentredString(x_centro, y_base, texto)
    
    def _dibujar_cuadro(self, c, x, y, codigo):
        """dibuja un cuadro completo con su codigo de barras"""
        # dibujar borde
        c.setLineWidth(1)
        c.setStrokeColorRGB(0, 0, 0)
        c.setFillColorRGB(0, 0, 0)
        c.rect(x, y, self.config.ANCHO_CUADRO, self.config.ALTO_CUADRO)
        
        # calcular centro del cuadro
        centro_x = x + (self.config.ANCHO_CUADRO / 2)
        
        # dibujar titulo del cuadro
        c.setFont(self.fuente, self.config.TAMANO_FUENTE_CUADRO)
        alto_titulo = 0.4 * cm
        y_titulo = y + self.config.ALTO_CUADRO - alto_titulo - 0.2 * cm
        c.drawCentredString(centro_x, y_titulo, self.config.TITULO_CUADRO)
        
        # calcular posiciones para codigo de barras
        altura_total_visual = 1.34 * cm
        espacio_texto = altura_total_visual - self.config.ALTO_BARRAS
        
        # calcular base del bloque visual
        y_base_bloque = y + (self.config.ALTO_CUADRO - altura_total_visual) / 2
        y_base_bloque += self.config.AJUSTE_VERTICAL_CODIGO
        
        # dibujar texto del codigo
        y_texto = y_base_bloque + 0.1 * cm
        self._dibujar_texto_codigo(c, centro_x, y_texto, codigo)
        
        # dibujar codigo de barras
        y_barras = y_base_bloque + espacio_texto + 0.03 * cm
        self._dibujar_codigo_barras(c, centro_x, y_barras, codigo)
    
    def _calcular_posiciones_x(self):
        """calcula las posiciones x de cada columna"""
        posiciones = []
        x_actual = self.config.MARGEN_IZQUIERDO
        
        for _ in range(self.config.COLUMNAS):
            posiciones.append(x_actual)
            x_actual += self.config.ANCHO_CUADRO + self.config.ESPACIO_HORIZONTAL
        
        return posiciones
    
    def _calcular_posiciones_y(self, alto_hoja):
        """calcula las posiciones y de cada fila"""
        posiciones = []
        y_actual = self.config.Y_INICIAL_GRID
        
        for _ in range(self.config.FILAS):
            # convertir desde arriba hacia abajo
            y_real = alto_hoja - y_actual - self.config.ALTO_CUADRO
            posiciones.append(y_real)
            y_actual += self.config.ALTO_CUADRO + self.config.ESPACIO_VERTICAL
        
        return posiciones
    
    def generar_pdf(self):
        """genera el pdf con todos los codigos de barras"""
        # crear canvas
        c = canvas.Canvas(self.config.NOMBRE_ARCHIVO, pagesize=A4)
        ancho_hoja, alto_hoja = A4
        
        # dibujar titulo principal
        self._dibujar_titulo_principal(c, ancho_hoja, alto_hoja)
        
        # calcular posiciones
        posiciones_x = self._calcular_posiciones_x()
        posiciones_y = self._calcular_posiciones_y(alto_hoja)
        
        # dibujar todos los cuadros
        contador = 0
        for y in posiciones_y:
            for x in posiciones_x:
                codigo = f"{self.config.CODIGO_INICIAL + contador}"
                self._dibujar_cuadro(c, x, y, codigo)
                contador += 1
        
        # guardar pdf
        c.showPage()
        c.save()
        print(f"generado correctamente: {self.config.NOMBRE_ARCHIVO}")
        print(f"total de etiquetas: {contador}")


# ==================== EJECUCION ====================
if __name__ == "__main__":
    generador = GeneradorEtiquetas()
    generador.generar_pdf()