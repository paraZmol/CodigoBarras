from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import cm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.graphics.barcode import code39
from reportlab.lib.utils import ImageReader
import openpyxl
import os


# ******************************************** zona de configuracion ********************************************

class Config:
    """configuracion centralizada del generador de etiquetas"""
    
    # archivo excel
    NOMBRE_EXCEL = "LIBROS FIIA.xlsx"
    FILA_INICIAL = 35
    FILA_FINAL = 101
    COLUMNA_CODIGOS = "L"        # codigo de barras
    COLUMNA_ESTANTERIA = "K"     # ubicacion en la estanteria
    
    # para el nombre de archivo
    NUMERO_ARCHIVO = "1"
    ABREVIACION_FACULTAD = "FIIA"
    
    # fuente
    RUTA_FUENTE = "OpenSans-Bold.ttf"
    RUTA_FUENTE_CODE = "OpenSans-Semibold.ttf"
    
    # argen de pagina
    MARGEN_SUPERIOR = 1.1 * cm
    MARGEN_IZQUIERDO = 0.4 * cm
    
    # --- grid de etiquetas ---
    FILAS = 8
    COLUMNAS = 3
    CUADROS_POR_HOJA = FILAS * COLUMNAS
    
    ESPACIO_HORIZONTAL = 0.15 * cm  # espacio entre columnas
    ESPACIO_VERTICAL = 0.06 * cm    # espacio entre filas
    Y_INICIAL_GRID = 2.70 * cm
    
    # dimensiones de cuadro individual 
    ANCHO_CUADRO = 6.56 * cm
    ALTO_CUADRO = 3.28 * cm
    
    # tipografia 
    TAMANO_FUENTE_TITULO = 13
    TAMANO_FUENTE_CUADRO = 10
    TAMANO_FUENTE_CODIGO = 8
    
    # codigo de barras visual 
    ANCHO_BARRAS = 0.05 * cm
    ALTO_BARRAS = 0.94 * cm
    MARGEN_HORIZONTAL_BARRAS = 0.1 * cm
    
    # codigo de barras textual
    MARGEN_HORIZONTAL_TEXTO = 1 * cm 
    SEPARACION_TEXTO_BARRAS = 0.3 * cm
    
    # ajustes finos
    AJUSTE_VERTICAL_CODIGO = 0.1 * cm  # ajuste vertical del bloque barras + texto
    
    # contenido del titulo
    TITULO_CUADRO = "Sistema de Bibliotecas UNASAM"


# ********************************************** lectura de datos **********************************************

class LectorExcel:
    """maneja la lectura del archivo excel"""
    
    def __init__(self, config):
        self.config = config
        self.workbook = None
        self.sheet = None
    
    def cargar_excel(self):
        """carga el archivo excel y retorna true si fue exitoso"""
        try:
            self.workbook = openpyxl.load_workbook(
                self.config.NOMBRE_EXCEL, 
                data_only=True
            )
            self.sheet = self.workbook.active
            print(f"excel cargado - {self.config.NOMBRE_EXCEL}")
            return True
        except FileNotFoundError:
            print(f"error - no se encontro '{self.config.NOMBRE_EXCEL}'")
            return False
        except Exception as e:
            print(f"error al cargar excel - {e}")
            return False
    
    def leer_codigos(self):
        """lee los codigos de barras del rango especificado"""
        codigos = []
        for fila in range(self.config.FILA_INICIAL, self.config.FILA_FINAL + 1):
            celda = f"{self.config.COLUMNA_CODIGOS}{fila}"
            valor = self.sheet[celda].value
            
            if valor is None or str(valor).strip() == "":
                codigos.append("*0*")
            else:
                codigos.append(str(valor).strip())
        
        print(f"codigos leidos - {len(codigos)}")
        return codigos
    
    def leer_rangos(self):
        """lee el rango inicial y final de estanteria"""
        celda_inicial = f"{self.config.COLUMNA_ESTANTERIA}{self.config.FILA_INICIAL}"
        celda_final = f"{self.config.COLUMNA_ESTANTERIA}{self.config.FILA_FINAL}"
        
        rango_inicial = self.sheet[celda_inicial].value
        rango_final = self.sheet[celda_final].value
        
        rango_inicial = str(rango_inicial).strip() if rango_inicial else ""
        rango_final = str(rango_final).strip() if rango_final else ""
        
        return rango_inicial, rango_final
    
    def cerrar(self):
        """cierra el archivo excel"""
        if self.workbook:
            self.workbook.close()


# ********************************************** generacion del pdf ********************************************** 

class GeneradorEtiquetas:
    """genera el pdf con las etiquetas de codigos de barras"""
    
    def __init__(self, config, rango_inicial, rango_final):
        self.config = config
        self.rango_inicial = rango_inicial
        self.rango_final = rango_final
        self.fuente_bold = None
        self.fuente_light = None
        self._cargar_fuentes()
    
    # inicializacion 
    
    def _cargar_fuentes(self):
        """carga las fuentes personalizadas o usa alternativas"""
        # fuente bold
        if os.path.exists(self.config.RUTA_FUENTE):
            pdfmetrics.registerFont(TTFont('OpenSans-Bold', self.config.RUTA_FUENTE))
            self.fuente_bold = "OpenSans-Bold"
        else:
            print(f"aviso - no se encontro '{self.config.RUTA_FUENTE}', usando helvetica-bold")
            self.fuente_bold = "Helvetica-Bold"
        
        # fuente light
        if os.path.exists(self.config.RUTA_FUENTE_CODE):
            pdfmetrics.registerFont(TTFont('OpenSans-Light', self.config.RUTA_FUENTE_CODE))
            self.fuente_light = "OpenSans-Light"
        else:
            print(f"aviso - no se encontro '{self.config.RUTA_FUENTE_CODE}', usando fuente principal")
            self.fuente_light = self.fuente_bold
    
    def _obtener_nombre_archivo(self):
        """genera el nombre del archivo pdf de salida"""
        rango_inicial_limpio = self.rango_inicial.replace('*', '').replace('/', '-').replace('\\', '-').replace(':', '-')
        rango_final_limpio = self.rango_final.replace('*', '').replace('/', '-').replace('\\', '-').replace(':', '-')
        return f"{self.config.NUMERO_ARCHIVO}{self.config.ABREVIACION_FACULTAD} {rango_inicial_limpio} - {rango_final_limpio}.pdf"
    
    # **************************** dibujo de titulos ****************************
    
    def _dibujar_titulo_principal(self, c, ancho_hoja, alto_hoja):
        """dibuja el titulo principal en la parte superior de la pagina"""
        pos_x = 5.52 * cm
        ancho = 9.97 * cm
        alto = 1.18 * cm
        pos_y = alto_hoja - self.config.MARGEN_SUPERIOR - alto
        centro_x = pos_x + (ancho / 2)
        
        c.setFillColorRGB(0, 0, 0)
        c.setFont(self.fuente_bold, self.config.TAMANO_FUENTE_TITULO)
        
        # linea 1 - biblioteca y ubicacion
        y_linea1 = pos_y + alto - 0.35 * cm
        c.drawCentredString(
            centro_x, 
            y_linea1, 
            f"BIBLIOTECA {self.config.ABREVIACION_FACULTAD} - UBICACION ESTANTERIA"
        )
        
        # linea 2 - rango de estanteria
        y_linea2 = pos_y + 0.25 * cm
        c.drawCentredString(
            centro_x, 
            y_linea2, 
            f"{self.rango_inicial} - {self.rango_final}"
        )
    
    # **************************** dibujo de elementos - codigo de barras ****************************
    
    def _dibujar_codigo_barras(self, c, x_cuadro, y_base, codigo):
        """dibuja el codigo de barras visual - barras negras"""
        codigo_limpio = codigo.replace("*", "")
        ancho_maximo = self.config.ANCHO_CUADRO - (2 * self.config.MARGEN_HORIZONTAL_BARRAS)
        
        # crear codigo de barras
        barcode = code39.Standard39(
            codigo_limpio,
            barHeight=self.config.ALTO_BARRAS,
            barWidth=self.config.ANCHO_BARRAS,
            checksum=0,
            humanReadable=False
        )
        
        # reducir si excede el ancho maximo
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
        
        # centrar horizontalmente en el cuadro
        centro_x_cuadro = x_cuadro + (self.config.ANCHO_CUADRO / 2)
        x_barcode = centro_x_cuadro - (barcode.width / 2)
        
        barcode.drawOn(c, x_barcode, y_base)
        return barcode.width
    
    def _dibujar_texto_codigo(self, c, x_cuadro, y_base, codigo):
        """dibuja el codigo en formato textual con justificacion expandida"""
        # zona de texto disponible
        margen = self.config.MARGEN_HORIZONTAL_TEXTO 
        ancho_util_texto = self.config.ANCHO_CUADRO - (2 * margen)
        x_inicio_texto = x_cuadro + margen
        
        c.setFont(self.fuente_light, self.config.TAMANO_FUENTE_CODIGO)
        
        # ajuste automatico del tamano de fuente si es necesario
        ancho_texto_puro = c.stringWidth(codigo, self.fuente_light, self.config.TAMANO_FUENTE_CODIGO)
        tamano_actual = self.config.TAMANO_FUENTE_CODIGO
        
        if ancho_texto_puro > ancho_util_texto:
            factor = ancho_util_texto / ancho_texto_puro
            tamano_actual = tamano_actual * factor
            c.setFont(self.fuente_light, tamano_actual)
        
        # justificacion expandida - letras espaciadas uniformemente
        num_caracteres = len(codigo)
        
        if num_caracteres <= 1:
            c.drawCentredString(x_inicio_texto + (ancho_util_texto / 2), y_base, codigo)
            return
        
        # calcular anchos individuales de cada letra
        ancho_solo_letras = 0
        anchos_individuales = []
        for letra in codigo:
            w = c.stringWidth(letra, self.fuente_light, tamano_actual)
            anchos_individuales.append(w)
            ancho_solo_letras += w
        
        # calcular espacio entre letras
        espacio_sobrante = ancho_util_texto - ancho_solo_letras
        if espacio_sobrante < 0:
            espacio_sobrante = 0
        
        gap = espacio_sobrante / (num_caracteres - 1)
        
        # dibujar cada letra con el espaciado calculado
        x_cursor = x_inicio_texto
        for i, letra in enumerate(codigo):
            c.drawString(x_cursor, y_base, letra)
            x_cursor += anchos_individuales[i] + gap
    
    # ************************************** dibujo de elementos - cuadro individual completo **************************************
    
    def _dibujar_cuadro(self, c, x, y, codigo):
        """dibuja un cuadro individual con titulo, codigo de barras y texto"""
        # borde del cuadro
        c.setLineWidth(1)
        c.setStrokeColorRGB(0, 0, 0)
        c.setFillColorRGB(0, 0, 0)
        c.rect(x, y, self.config.ANCHO_CUADRO, self.config.ALTO_CUADRO)
        
        centro_x = x + (self.config.ANCHO_CUADRO / 2)
        
        # titulo interno del cuadro
        c.setFont(self.fuente_bold, self.config.TAMANO_FUENTE_CUADRO)
        alto_titulo = 0.4 * cm
        y_titulo = y + self.config.ALTO_CUADRO - alto_titulo - 0.2 * cm
        c.drawCentredString(centro_x, y_titulo, self.config.TITULO_CUADRO)
        
        # calculos para el bloque de codigo - barras + texto
        altura_total_visual = 1.34 * cm
        espacio_texto_total = altura_total_visual - self.config.ALTO_BARRAS
        
        y_base_bloque = y + (self.config.ALTO_CUADRO - altura_total_visual) / 2
        y_base_bloque += self.config.AJUSTE_VERTICAL_CODIGO  # ajuste fino
        
        # dibujar codigo de barras visual
        y_barras = y_base_bloque + espacio_texto_total + 0.03 * cm
        self._dibujar_codigo_barras(c, x, y_barras, codigo)
        
        # dibujar codigo textual
        y_texto = y_barras - self.config.SEPARACION_TEXTO_BARRAS
        self._dibujar_texto_codigo(c, x, y_texto, codigo)
    
    # *************************************** calculo de posiciones ***************************************
    
    def _calcular_posiciones_x(self):
        """calcula las posiciones x de las columnas del grid"""
        posiciones = []
        x_actual = self.config.MARGEN_IZQUIERDO
        
        for _ in range(self.config.COLUMNAS):
            posiciones.append(x_actual)
            x_actual += self.config.ANCHO_CUADRO + self.config.ESPACIO_HORIZONTAL
        
        return posiciones
    
    def _calcular_posiciones_y(self, alto_hoja):
        """calcula las posiciones y de las filas del grid"""
        posiciones = []
        y_actual = self.config.Y_INICIAL_GRID
        
        for _ in range(self.config.FILAS):
            y_real = alto_hoja - y_actual - self.config.ALTO_CUADRO
            posiciones.append(y_real)
            y_actual += self.config.ALTO_CUADRO + self.config.ESPACIO_VERTICAL
        
        return posiciones
    
    # ******************************** composicion de pagina ********************************
    
    def _dibujar_pagina(self, c, codigos_pagina, ancho_hoja, alto_hoja):
        """dibuja una pagina completa con titulo y grid de etiquetas"""
        # titulo principal
        self._dibujar_titulo_principal(c, ancho_hoja, alto_hoja)
        
        # calcular posiciones del grid
        posiciones_x = self._calcular_posiciones_x()
        posiciones_y = self._calcular_posiciones_y(alto_hoja)
        
        # dibujar cada cuadro en el grid
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
    
    # **************************** generacion principal ****************************
    
    def generar_pdf(self, codigos):
        """genera el archivo pdf completo con todas las etiquetas"""
        nombre_archivo = self._obtener_nombre_archivo()
        c = canvas.Canvas(nombre_archivo, pagesize=A4)
        ancho_hoja, alto_hoja = A4
        
        total_codigos = len(codigos)
        total_paginas = (total_codigos + self.config.CUADROS_POR_HOJA - 1) // self.config.CUADROS_POR_HOJA
        
        print(f"\n{'=' * 50}")
        print(f"generando {total_paginas} pagina-s-...")
        print(f"{'=' * 50}\n")
        
        # generar cada pagina
        for num_pagina in range(total_paginas):
            inicio = num_pagina * self.config.CUADROS_POR_HOJA
            fin = min(inicio + self.config.CUADROS_POR_HOJA, total_codigos)
            codigos_pagina = codigos[inicio:fin]
            self._dibujar_pagina(c, codigos_pagina, ancho_hoja, alto_hoja)
            c.showPage()
        
        c.save()
        print(f"\ngenerado correctamente - {nombre_archivo}")
        print(f"total de etiquetas - {total_codigos}")


# ************************************* ejecucion principal *************************************

def main():
    """funcion principal que ejecuta todo el proceso"""
    config = Config()
    lector = LectorExcel(config)
    
    # cargar excel
    if not lector.cargar_excel(): 
        return
    
    # leer datos
    codigos = lector.leer_codigos()
    rango_inicial, rango_final = lector.leer_rangos()
    lector.cerrar()
    
    # validar que haya codigos
    if not codigos:
        print("error - no hay codigos")
        return
    
    # generar pdf
    generador = GeneradorEtiquetas(config, rango_inicial, rango_final)
    generador.generar_pdf(codigos)


if __name__ == "__main__":
    main()