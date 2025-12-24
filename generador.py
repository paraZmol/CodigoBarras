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
    FILA_INICIAL = 300
    FILA_FINAL = 370
    COLUMNA_CODIGOS = "L"        # codigo de barras
    COLUMNA_ESTANTERIA = "K"     # ubicacion en la estanteria
    
    # para el nombre de archivo (este valor sera ignorado por el autoincrementable)
    # NUMERO_ARCHIVO = "5" 
    ABREVIACION_FACULTAD = "FIIA"
    
    # fuentes
    RUTA_FUENTE = "OpenSans-Bold.ttf"
    RUTA_FUENTE_CODE = "OpenSans-Semibold.ttf"
    
    # configuracion de imagenes - logos
    RUTA_LOGO_UNASAM = "logo-unasam.png"
    RUTA_LOGO_FACULTAD = "facultad.png"
    
    # dimensiones y ubicacion de imagenes
    ALTO_IMAGENES = 0.7 * cm
    
    # margenes horizontales para logos
    MARGEN_X_LOGO_UNASAM = 0.7 * cm      # margen izquierdo para logo unasam
    MARGEN_X_LOGO_FACULTAD = 0.7 * cm    # margen derecho para logo facultad
    
    # distancia vertical de las imagenes desde el codigo
    DISTANCIA_Y_DESDE_CODIGO = -1.3 * cm
    
    # margen de pagina
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
        self.fuente_code = None
        self._cargar_fuentes()
        
        # logica de autoincremento al iniciar
        self.numero_archivo_auto = self._calcular_siguiente_numero()
    
    # inicializacion 
    
    def _cargar_fuentes(self):
        """carga las fuentes personalizadas o usa alternativas"""
        # fuente bold (titulos)
        if os.path.exists(self.config.RUTA_FUENTE):
            pdfmetrics.registerFont(TTFont('OpenSans-Bold', self.config.RUTA_FUENTE))
            self.fuente_bold = "OpenSans-Bold"
        else:
            print(f"aviso - no se encontro '{self.config.RUTA_FUENTE}', usando helvetica-bold")
            self.fuente_bold = "Helvetica-Bold"
        
        # fuente code (texto del codigo)
        if os.path.exists(self.config.RUTA_FUENTE_CODE):
            # registramos con un nombre interno claro 'OpenSans-Code'
            pdfmetrics.registerFont(TTFont('OpenSans-Code', self.config.RUTA_FUENTE_CODE))
            self.fuente_code = "OpenSans-Code"
        else:
            print(f"aviso - no se encontro '{self.config.RUTA_FUENTE_CODE}', usando fuente principal")
            self.fuente_code = self.fuente_bold

    def _calcular_siguiente_numero(self):
        """
        busca el siguiente numero de archivo basado en lo que existe en la carpeta
        logica: busca archivos que empiecen con numero y sigan con la facultad
        ejemplo: '5FIIA...' -> detecta 5
        """
        facultad = self.config.ABREVIACION_FACULTAD
        archivos = [f for f in os.listdir('.') if f.endswith('.pdf')]
        
        max_numero = 0
        
        print(f"buscando archivos anteriores de {facultad}...")
        
        for archivo in archivos:
            # verificar si contiene la abreviacion de facultad
            if facultad in archivo:
                # dividir el nombre en la primera aparicion de la facultad
                # ej: "5FIIA 500-600.pdf" -> ["5", " 500-600.pdf"]
                partes = archivo.split(facultad, 1)
                
                if len(partes) > 1:
                    prefijo = partes[0]
                    
                    # verificar estrictamente si es un numero entero lo que esta antes
                    if prefijo.isdigit():
                        numero = int(prefijo)
                        if numero > max_numero:
                            max_numero = numero
        
        siguiente = max_numero + 1
        print(f"ultimo encontrado: {max_numero if max_numero > 0 else 'ninguno'} -> generando archivo numero: {siguiente}")
        return str(siguiente)
    
    def _obtener_nombre_archivo(self):
        """genera el nombre del archivo pdf usando el numero calculado"""
        rango_inicial_limpio = self.rango_inicial.replace('*', '').replace('/', '-').replace('\\', '-').replace(':', '-')
        rango_final_limpio = self.rango_final.replace('*', '').replace('/', '-').replace('\\', '-').replace(':', '-')
        
        # usa self.numero_archivo_auto
        return f"{self.numero_archivo_auto}{self.config.ABREVIACION_FACULTAD} {rango_inicial_limpio} - {rango_final_limpio}.pdf"
    
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
    
    # **************************** dibujo de elementos - imagenes ****************************
    
    def _dibujar_imagen(self, c, ruta_imagen, x, y, alto_deseado):
        """dibuja una imagen redimensionada proporcionalmente"""
        if not os.path.exists(ruta_imagen):
            print(f"aviso - imagen no encontrada '{ruta_imagen}'")
            return 0

        try:
            # obtener dimensiones reales
            img_utils = ImageReader(ruta_imagen)
            ancho_real, alto_real = img_utils.getSize()
            
            # calcular aspecto - ancho / alto
            aspect_ratio = ancho_real / alto_real
            
            # calcular nuevo ancho basado en el alto deseado
            nuevo_ancho = alto_deseado * aspect_ratio
            
            # dibujar imagen - reportlab dibuja desde esquina inferior izquierda
            c.drawImage(ruta_imagen, x, y, width=nuevo_ancho, height=alto_deseado, mask='auto')
            
            return nuevo_ancho
            
        except Exception as e:
            print(f"error al dibujar imagen {ruta_imagen} - {e}")
            return 0
            
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
        
        c.setFont(self.fuente_code, self.config.TAMANO_FUENTE_CODIGO)
        
        # ajuste automatico del tamano de fuente si es necesario
        ancho_texto_puro = c.stringWidth(codigo, self.fuente_code, self.config.TAMANO_FUENTE_CODIGO)
        tamano_actual = self.config.TAMANO_FUENTE_CODIGO
        
        if ancho_texto_puro > ancho_util_texto:
            factor = ancho_util_texto / ancho_texto_puro
            tamano_actual = tamano_actual * factor
            c.setFont(self.fuente_code, tamano_actual)
        
        # justificacion expandida - letras espaciadas uniformemente
        num_caracteres = len(codigo)
        
        if num_caracteres <= 1:
            c.drawCentredString(x_inicio_texto + (ancho_util_texto / 2), y_base, codigo)
            return
        
        # calcular anchos individuales de cada letra
        ancho_solo_letras = 0
        anchos_individuales = []
        for letra in codigo:
            w = c.stringWidth(letra, self.fuente_code, tamano_actual)
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
        
        # posiciones verticales calculadas
        y_barras = y_base_bloque + espacio_texto_total + 0.03 * cm
        y_texto = y_barras - self.config.SEPARACION_TEXTO_BARRAS
        y_imagenes = y_barras + self.config.DISTANCIA_Y_DESDE_CODIGO
        
        # --- DIBUJO DE IMAGENES ---
        
        # 1. dibujar logo unasam - izquierda
        x_logo_unasam = x + self.config.MARGEN_X_LOGO_UNASAM
        self._dibujar_imagen(c, self.config.RUTA_LOGO_UNASAM, x_logo_unasam, y_imagenes, self.config.ALTO_IMAGENES)
        
        # 2. dibujar logo facultad - derecha
        # primero obtenemos el ancho real que tendra la imagen
        ancho_img_facultad = 0
        if os.path.exists(self.config.RUTA_LOGO_FACULTAD):
            try:
                img_fac = ImageReader(self.config.RUTA_LOGO_FACULTAD)
                w_f, h_f = img_fac.getSize()
                aspect_f = w_f / h_f
                ancho_img_facultad = self.config.ALTO_IMAGENES * aspect_f
            except:
                ancho_img_facultad = 0 # fallo lectura
        
        if ancho_img_facultad > 0:
            # calculo de x derecha - ancho cuadro - margen derecho - ancho imagen
            x_logo_facultad = (x + self.config.ANCHO_CUADRO) - self.config.MARGEN_X_LOGO_FACULTAD - ancho_img_facultad
            self._dibujar_imagen(c, self.config.RUTA_LOGO_FACULTAD, x_logo_facultad, y_imagenes, self.config.ALTO_IMAGENES)
        
        # --- DIBUJO DE CODIGOS ---
        
        # dibujar codigo de barras visual
        self._dibujar_codigo_barras(c, x, y_barras, codigo)
        
        # dibujar codigo textual
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