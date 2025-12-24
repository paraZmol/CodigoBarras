from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.barcode import code39
import openpyxl
import os

# ******************* ZONA DE CONFIGURACION *******************
class Config:
    # archivo excel de entrada
    NOMBRE_EXCEL = "LIBROS FIIA.xlsx"
    
    # rango de filas en excel
    FILA_INICIAL = 35
    FILA_FINAL = 101
    
    COLUMNA_CODIGOS = "L" # codigo de barras
    COLUMNA_ESTANTERIA = "K" # ubicacion en la estanteria
    
    # configuracion para nombre de archivo y titulo
    NUMERO_ARCHIVO = "1"
    ABREVIACION_FACULTAD = "FIIA"
    
    # fuentes personalizadas
    RUTA_FUENTE = "OpenSans-Bold.ttf"
    RUTA_FUENTE_CODE = "OpenSans-Semibold.ttf"
    
    # margenes principales del papel
    MARGEN_SUPERIOR = 1.1 * cm
    MARGEN_IZQUIERDO = 0.4 * cm
    
    # espaciado entre cuadros - grid
    ESPACIO_HORIZONTAL = 0.15 * cm  # espacio entre columnas
    ESPACIO_VERTICAL = 0.06 * cm    # espacio entre filas
    
    # posicion inicial del grid
    Y_INICIAL_GRID = 2.70 * cm
    
    # dimensiones del cuadro individual
    ANCHO_CUADRO = 6.56 * cm
    ALTO_CUADRO = 3.28 * cm
    
    # ajuste fino vertical del bloque completo *** barras + texto
    AJUSTE_VERTICAL_CODIGO = 0.1 * cm
    
    # cantidad de cuadros por hoja
    FILAS = 8
    COLUMNAS = 3
    CUADROS_POR_HOJA = FILAS * COLUMNAS
    
    # configuracion del titulo principal
    TAMANO_FUENTE_TITULO = 13
    
    # configuracion del cuadro individual
    TITULO_CUADRO = "Sistema de Bibliotecas UNASAM"
    TAMANO_FUENTE_CUADRO = 10
    TAMANO_FUENTE_CODIGO = 8
    
    # configuracion del codigo de barras visual 
    ANCHO_BARRAS = 0.05 * cm
    ALTO_BARRAS = 0.94 * cm
    
    # margenes laterales para el codigo de barras VISUAL
    MARGEN_HORIZONTAL_BARRAS = 0.1 * cm
    
    # margen lateral para el codigo TEXTUAL
    MARGEN_HORIZONTAL_TEXTO = 1 * cm

    # margen entre codigos
    SEPARACION_TEXTO_BARRAS = 0.3 * cm

# *************************************************************


class LectorExcel:
    """maneja la lectura del archivo excel"""
    
    def __init__(self, config):
        self.config = config
        self.workbook = None
        self.sheet = None
    
    def cargar_excel(self):
        try:
            self.workbook = openpyxl.load_workbook(self.config.NOMBRE_EXCEL, data_only=True)
            self.sheet = self.workbook.active
            print(f"excel cargado: {self.config.NOMBRE_EXCEL}")
            return True
        except FileNotFoundError:
            print(f"error: no se encontro el archivo '{self.config.NOMBRE_EXCEL}'")
            return False
        except Exception as e:
            print(f"error al cargar excel: {e}")
            return False
    
    def leer_codigos(self):
        codigos = []
        for fila in range(self.config.FILA_INICIAL, self.config.FILA_FINAL + 1):
            celda = f"{self.config.COLUMNA_CODIGOS}{fila}"
            valor = self.sheet[celda].value
            if valor is None or str(valor).strip() == "":
                codigos.append("*0*")
            else:
                codigos.append(str(valor).strip())
        print(f"codigos leidos: {len(codigos)}")
        return codigos
    
    def leer_rangos(self):
        celda_inicial = f"{self.config.COLUMNA_ESTANTERIA}{self.config.FILA_INICIAL}"
        celda_final = f"{self.config.COLUMNA_ESTANTERIA}{self.config.FILA_FINAL}"
        rango_inicial = self.sheet[celda_inicial].value
        rango_final = self.sheet[celda_final].value
        rango_inicial = str(rango_inicial).strip() if rango_inicial else ""
        rango_final = str(rango_final).strip() if rango_final else ""
        return rango_inicial, rango_final
    
    def cerrar(self):
        if self.workbook:
            self.workbook.close()


class GeneradorEtiquetas:
    """genera el pdf con los codigos de barras"""
    
    def __init__(self, config, rango_inicial, rango_final):
        self.config = config
        self.rango_inicial = rango_inicial
        self.rango_final = rango_final
        self._cargar_fuentes()
    
    def _cargar_fuentes(self):
        """Carga las fuentes Bold y Light para usarlas separadamente"""
        if os.path.exists(self.config.RUTA_FUENTE):
            pdfmetrics.registerFont(TTFont('OpenSans-Bold', self.config.RUTA_FUENTE))
            self.fuente_bold = "OpenSans-Bold"
        else:
            print(f"aviso: no se encontro '{self.config.RUTA_FUENTE}', usando helvetica")
            self.fuente_bold = "Helvetica-Bold"
            
        if os.path.exists(self.config.RUTA_FUENTE_CODE):
            pdfmetrics.registerFont(TTFont('OpenSans-Light', self.config.RUTA_FUENTE_CODE))
            self.fuente_light = "OpenSans-Light"
        else:
            print(f"aviso: no se encontro '{self.config.RUTA_FUENTE_CODE}', usando fuente principal")
            self.fuente_light = self.fuente_bold

    def _obtener_nombre_archivo(self):
        rango_inicial_limpio = self.rango_inicial.replace('*', '').replace('/', '-').replace('\\', '-').replace(':', '-')
        rango_final_limpio = self.rango_final.replace('*', '').replace('/', '-').replace('\\', '-').replace(':', '-')
        return f"{self.config.NUMERO_ARCHIVO}{self.config.ABREVIACION_FACULTAD} {rango_inicial_limpio} - {rango_final_limpio}.pdf"
    
    def _dibujar_titulo_principal(self, c, ancho_hoja, alto_hoja):
        pos_x = 5.52 * cm
        ancho = 9.97 * cm
        alto = 1.18 * cm
        pos_y = alto_hoja - self.config.MARGEN_SUPERIOR - alto
        centro_x = pos_x + (ancho / 2)
        
        c.setFillColorRGB(0, 0, 0)
        c.setFont(self.fuente_bold, self.config.TAMANO_FUENTE_TITULO)
        
        y_linea1 = pos_y + alto - 0.35 * cm
        c.drawCentredString(centro_x, y_linea1, f"BIBLIOTECA {self.config.ABREVIACION_FACULTAD} - UBICACIÓN ESTANTERIA")
        
        y_linea2 = pos_y + 0.25 * cm
        c.drawCentredString(centro_x, y_linea2, f"{self.rango_inicial} - {self.rango_final}")
    
    def _dibujar_codigo_barras(self, c, x_cuadro, y_base, codigo):
        codigo_limpio = codigo.replace("*", "")
        ancho_maximo = self.config.ANCHO_CUADRO - (2 * self.config.MARGEN_HORIZONTAL_BARRAS)
        
        barcode = code39.Standard39(
            codigo_limpio,
            barHeight=self.config.ALTO_BARRAS,
            barWidth=self.config.ANCHO_BARRAS,
            checksum=0,
            humanReadable=False
        )
        
        if barcode.width > ancho_maximo:
            factor_reduccion = ancho_maximo / barcode.width
            nuevo_ancho_barra = self.config.ANCHO_BARRAS * factor_reduccion
            barcode = code39.Standard39(
                codigo_limpio,
                barHeight=self.config.ALTO_BARRAS,
                barWidth=nuevo_ancho_barra,
                checksum=0,
                humanReadable=False
            )
        
        centro_x_cuadro = x_cuadro + (self.config.ANCHO_CUADRO / 2)
        x_barcode = centro_x_cuadro - (barcode.width / 2)
        barcode.drawOn(c, x_barcode, y_base)
        return barcode.width
    
    def _dibujar_texto_codigo(self, c, x_cuadro, y_base, codigo):
        # zona de texto
        margen = self.config.MARGEN_HORIZONTAL_TEXTO
        ancho_util_texto = self.config.ANCHO_CUADRO - (2 * margen)
        x_inicio_texto = x_cuadro + margen
        
        c.setFont(self.fuente_light, self.config.TAMANO_FUENTE_CODIGO)
        
        # reajuste del tamaño de la fuente
        ancho_texto_puro = c.stringWidth(codigo, self.fuente_light, self.config.TAMANO_FUENTE_CODIGO)
        tamano_actual = self.config.TAMANO_FUENTE_CODIGO
        if ancho_texto_puro > ancho_util_texto:
            factor = ancho_util_texto / ancho_texto_puro
            tamano_actual = tamano_actual * factor
            c.setFont(self.fuente_light, tamano_actual)
            
        # justificado de letras
        num_caracteres = len(codigo)
        if num_caracteres <= 1:
            c.drawCentredString(x_inicio_texto + (ancho_util_texto/2), y_base, codigo)
            return

        ancho_solo_letras = 0
        anchos_individuales = []
        for letra in codigo:
            w = c.stringWidth(letra, self.fuente_light, tamano_actual)
            anchos_individuales.append(w)
            ancho_solo_letras += w
            
        espacio_sobrante = ancho_util_texto - ancho_solo_letras
        if espacio_sobrante < 0: espacio_sobrante = 0
        
        gap = espacio_sobrante / (num_caracteres - 1)
        
        x_cursor = x_inicio_texto
        for i, letra in enumerate(codigo):
            c.drawString(x_cursor, y_base, letra)
            x_cursor += anchos_individuales[i] + gap

    def _dibujar_cuadro(self, c, x, y, codigo):
        # borde
        c.setLineWidth(1)
        c.setStrokeColorRGB(0, 0, 0)
        c.setFillColorRGB(0, 0, 0)
        c.rect(x, y, self.config.ANCHO_CUADRO, self.config.ALTO_CUADRO)
        
        centro_x = x + (self.config.ANCHO_CUADRO / 2)
        
        # titulo interno
        c.setFont(self.fuente_bold, self.config.TAMANO_FUENTE_CUADRO)
        alto_titulo = 0.4 * cm
        y_titulo = y + self.config.ALTO_CUADRO - alto_titulo - 0.2 * cm
        c.drawCentredString(centro_x, y_titulo, self.config.TITULO_CUADRO)
        
        # calculos verticales generales
        altura_total_visual = 1.34 * cm
        espacio_texto_total = altura_total_visual - self.config.ALTO_BARRAS
        
        # base del bloque
        y_base_bloque = y + (self.config.ALTO_CUADRO - altura_total_visual) / 2
        y_base_bloque += self.config.AJUSTE_VERTICAL_CODIGO
        
        # dibujo de barras
        y_barras = y_base_bloque + espacio_texto_total + 0.03 * cm
        self._dibujar_codigo_barras(c, x, y_barras, codigo)
        
        # deibujo de texto
        y_texto = y_barras - self.config.SEPARACION_TEXTO_BARRAS
        self._dibujar_texto_codigo(c, x, y_texto, codigo)
    
    def _calcular_posiciones_x(self):
        posiciones = []
        x_actual = self.config.MARGEN_IZQUIERDO
        for _ in range(self.config.COLUMNAS):
            posiciones.append(x_actual)
            x_actual += self.config.ANCHO_CUADRO + self.config.ESPACIO_HORIZONTAL
        return posiciones
    
    def _calcular_posiciones_y(self, alto_hoja):
        posiciones = []
        y_actual = self.config.Y_INICIAL_GRID
        for _ in range(self.config.FILAS):
            y_real = alto_hoja - y_actual - self.config.ALTO_CUADRO
            posiciones.append(y_real)
            y_actual += self.config.ALTO_CUADRO + self.config.ESPACIO_VERTICAL
        return posiciones
    
    def _dibujar_pagina(self, c, codigos_pagina, ancho_hoja, alto_hoja):
        self._dibujar_titulo_principal(c, ancho_hoja, alto_hoja)
        posiciones_x = self._calcular_posiciones_x()
        posiciones_y = self._calcular_posiciones_y(alto_hoja)
        
        indice = 0
        for y in posiciones_y:
            for x in posiciones_x:
                if indice < len(codigos_pagina):
                    self._dibujar_cuadro(c, x, y, codigos_pagina[indice])
                    indice += 1
                else:
                    break
            if indice >= len(codigos_pagina):
                break
    
    def generar_pdf(self, codigos):
        nombre_archivo = self._obtener_nombre_archivo()
        c = canvas.Canvas(nombre_archivo, pagesize=A4)
        ancho_hoja, alto_hoja = A4
        
        total_codigos = len(codigos)
        total_paginas = (total_codigos + self.config.CUADROS_POR_HOJA - 1) // self.config.CUADROS_POR_HOJA
        
        print(f"generando {total_paginas} pagina(s)...")
        
        for num_pagina in range(total_paginas):
            inicio = num_pagina * self.config.CUADROS_POR_HOJA
            fin = min(inicio + self.config.CUADROS_POR_HOJA, total_codigos)
            codigos_pagina = codigos[inicio:fin]
            self._dibujar_pagina(c, codigos_pagina, ancho_hoja, alto_hoja)
            c.showPage()
        
        c.save()
        print(f"\ngenerado correctamente: {nombre_archivo}")
        print(f"total de etiquetas: {total_codigos}")

# ******************************************** EJECUCION ********************************************
def main():
    config = Config()
    lector = LectorExcel(config)
    if not lector.cargar_excel(): return
    
    codigos = lector.leer_codigos()
    rango_inicial, rango_final = lector.leer_rangos()
    lector.cerrar()
    
    if not codigos:
        print("error: no hay codigos")
        return
    
    generador = GeneradorEtiquetas(config, rango_inicial, rango_final)
    generador.generar_pdf(codigos)

if __name__ == "__main__":
    main()