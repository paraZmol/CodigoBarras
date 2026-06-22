"""
Interfaz grafica para el generador de etiquetas de codigos de barras.

Permite editar visualmente todos los parametros de configuracion (que antes
solo se podian cambiar en el codigo), ver una vista previa en vivo de como
queda una hoja de etiquetas, y generar los PDFs / pintar el Excel sin tocar
una sola linea de codigo.

Ejecutar con:  python interfaz.py
"""

import os
import sys
import tempfile
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from reportlab.lib.units import cm
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

import fitz  # PyMuPDF - para renderizar el PDF a imagen
from PIL import Image, ImageTk

# importamos la logica original sin modificarla
import generador
from generador import Config, GeneradorEtiquetas


# ============================================================================
#  Definicion de los campos editables
#  Cada campo: (atributo_en_Config, etiqueta_visible, tipo, en_cm)
#    tipo: "text" | "int" | "float"
#    en_cm: True  -> el valor mostrado esta en cm y se multiplica por `cm`
#                    al escribirlo en Config (Config guarda en puntos).
#           False -> el valor se usa tal cual.
# ============================================================================

SECCIONES = {
    "Archivos y datos": [
        ("NOMBRE_EXCEL", "Excel de entrada", "text", False),
        ("NOMBRE_EXCEL_SALIDA", "Excel de salida (pintado)", "text", False),
        ("FILA_INICIAL", "Fila inicial", "int", False),
        ("COLUMNA_CODIGOS", "Columna de codigos", "text", False),
        ("ABREVIACION_FACULTAD", "Sigla / facultad", "text", False),
    ],
    "Logos": [
        ("RUTA_LOGO_UNASAM", "Logo UNASAM", "text", False),
        ("RUTA_LOGO_FACULTAD", "Logo facultad", "text", False),
        ("ALTO_IMAGENES", "Alto de logos (cm)", "float", True),
        ("MARGEN_X_LOGO_UNASAM", "Margen X logo UNASAM (cm)", "float", True),
        ("MARGEN_X_LOGO_FACULTAD", "Margen X logo facultad (cm)", "float", True),
        ("DISTANCIA_Y_DESDE_CODIGO", "Distancia Y logos desde codigo (cm)", "float", True),
    ],
    "Titulos y textos": [
        ("TITULO_CUADRO", "Titulo del cuadro", "text", False),
        ("TAMANO_FUENTE_TITULO", "Tamano fuente titulo", "int", False),
        ("TAMANO_FUENTE_CUADRO", "Tamano fuente cuadro", "int", False),
        ("TAMANO_FUENTE_CODIGO", "Tamano fuente codigo", "int", False),
    ],
    "Margenes y grid": [
        ("MARGEN_SUPERIOR", "Margen superior (cm)", "float", True),
        ("MARGEN_IZQUIERDO", "Margen izquierdo (cm)", "float", True),
        ("FILAS", "Filas por hoja", "int", False),
        ("COLUMNAS", "Columnas por hoja", "int", False),
        ("ESPACIO_HORIZONTAL", "Espacio entre columnas (cm)", "float", True),
        ("ESPACIO_VERTICAL", "Espacio entre filas (cm)", "float", True),
        ("Y_INICIAL_GRID", "Y inicial del grid (cm)", "float", True),
    ],
    "Dimensiones del cuadro": [
        ("ANCHO_CUADRO", "Ancho cuadro (cm)", "float", True),
        ("ALTO_CUADRO", "Alto cuadro (cm)", "float", True),
    ],
    "Codigo de barras": [
        ("ANCHO_BARRAS", "Ancho de barras (cm)", "float", True),
        ("ALTO_BARRAS", "Alto de barras (cm)", "float", True),
        ("MARGEN_HORIZONTAL_BARRAS", "Margen horizontal barras (cm)", "float", True),
        ("MARGEN_HORIZONTAL_TEXTO", "Margen horizontal texto (cm)", "float", True),
        ("SEPARACION_TEXTO_BARRAS", "Separacion texto-barras (cm)", "float", True),
        ("AJUSTE_VERTICAL_CODIGO", "Ajuste vertical codigo (cm)", "float", True),
    ],
}


class InterfazGenerador:
    """Pestana 'Desde Excel' - la logica original, sin cambios funcionales.

    Antes ocupaba toda la ventana; ahora se monta dentro de un frame de
    pestana (`parent`). Todo lo demas se conserva igual.
    """

    def __init__(self, parent):
        self.root = parent  # frame de la pestana (compatibilidad de nombres)

        # base de directorio = donde vive el proyecto (para rutas relativas)
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        os.chdir(self.base_dir)

        self.vars = {}          # atributo -> tk.StringVar
        self.meta = {}          # atributo -> (tipo, en_cm)
        self._preview_imgtk = None  # mantener referencia viva de la imagen

        self._construir_layout()
        self._cargar_valores_por_defecto()
        # vista previa inicial
        self.root.after(300, self.actualizar_preview)

    # -------------------------------------------------------------- layout
    def _construir_layout(self):
        # barra inferior de acciones (se empaqueta primero para que no la tape
        # el contenedor expandible)
        barra = ttk.Frame(self.root)
        barra.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Separator(barra).pack(fill=tk.X)
        acciones = ttk.Frame(barra)
        acciones.pack(fill=tk.X, padx=8, pady=8)

        self.status = ttk.Label(acciones, text="Listo.", foreground="#444")
        self.status.pack(side=tk.LEFT)

        ttk.Button(acciones, text="Generar PDFs + Pintar Excel",
                   command=self.generar).pack(side=tk.RIGHT, padx=4)
        ttk.Button(acciones, text="Restaurar valores por defecto",
                   command=self._cargar_valores_por_defecto).pack(side=tk.RIGHT, padx=4)

        # panel izquierdo (formulario con scroll) | panel derecho (preview)
        contenedor = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        contenedor.pack(fill=tk.BOTH, expand=True)

        # ----- panel izquierdo con scroll -----
        izq = ttk.Frame(contenedor, width=440)
        contenedor.add(izq, weight=0)

        canvas_form = tk.Canvas(izq, borderwidth=0, width=440, highlightthickness=0)
        scroll = ttk.Scrollbar(izq, orient="vertical", command=canvas_form.yview)
        self.form_frame = ttk.Frame(canvas_form)

        self.form_frame.bind(
            "<Configure>",
            lambda e: canvas_form.configure(scrollregion=canvas_form.bbox("all")),
        )
        canvas_form.create_window((0, 0), window=self.form_frame, anchor="nw")
        canvas_form.configure(yscrollcommand=scroll.set)

        canvas_form.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        # scroll con la rueda del raton
        def _on_wheel(event):
            canvas_form.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas_form.bind_all("<MouseWheel>", _on_wheel)

        self._construir_campos()

        # ----- panel derecho: preview -----
        der = ttk.Frame(contenedor)
        contenedor.add(der, weight=1)

        cab = ttk.Frame(der)
        cab.pack(fill=tk.X, padx=8, pady=6)
        ttk.Label(cab, text="Vista previa (1 hoja)", font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT)
        ttk.Button(cab, text="Actualizar vista previa", command=self.actualizar_preview).pack(side=tk.RIGHT)

        self.preview_canvas = tk.Canvas(der, bg="#d9d9d9", highlightthickness=0)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 6))
        self.preview_canvas.bind("<Configure>", lambda e: self._redibujar_preview())

    def _construir_campos(self):
        for seccion, campos in SECCIONES.items():
            lf = ttk.LabelFrame(self.form_frame, text=seccion)
            lf.pack(fill=tk.X, padx=8, pady=6, anchor="n")

            for attr, etiqueta, tipo, en_cm in campos:
                fila = ttk.Frame(lf)
                fila.pack(fill=tk.X, padx=6, pady=3)

                ttk.Label(fila, text=etiqueta, width=30, anchor="w").pack(side=tk.LEFT)
                var = tk.StringVar()
                self.vars[attr] = var
                self.meta[attr] = (tipo, en_cm)

                # actualizar preview al perder foco / Enter
                entry = ttk.Entry(fila, textvariable=var)
                entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
                entry.bind("<Return>", lambda e: self.actualizar_preview())
                entry.bind("<FocusOut>", lambda e: self.actualizar_preview())

                # boton de buscar archivo para rutas
                if attr in ("NOMBRE_EXCEL", "RUTA_LOGO_UNASAM", "RUTA_LOGO_FACULTAD"):
                    ttk.Button(fila, text="...", width=3,
                               command=lambda v=var, a=attr: self._buscar_archivo(v, a)).pack(side=tk.LEFT, padx=(4, 0))

    def _buscar_archivo(self, var, attr):
        if attr == "NOMBRE_EXCEL":
            tipos = [("Excel", "*.xlsx *.xls"), ("Todos", "*.*")]
        else:
            tipos = [("Imagenes", "*.png *.jpg *.jpeg"), ("Todos", "*.*")]
        ruta = filedialog.askopenfilename(initialdir=self.base_dir, filetypes=tipos)
        if ruta:
            # guardar relativo si esta dentro del proyecto
            try:
                rel = os.path.relpath(ruta, self.base_dir)
                var.set(rel if not rel.startswith("..") else ruta)
            except ValueError:
                var.set(ruta)
            self.actualizar_preview()

    # ------------------------------------------------------- valores
    def _cargar_valores_por_defecto(self):
        base = Config()
        for attr, (tipo, en_cm) in self.meta.items():
            valor = getattr(base, attr)
            if en_cm:
                valor = round(valor / cm, 4)
            self.vars[attr].set(str(valor))
        self.actualizar_preview()

    def _construir_config(self):
        """crea una instancia de Config con los valores de la interfaz."""
        cfg = Config()
        for attr, var in self.vars.items():
            tipo, en_cm = self.meta[attr]
            txt = var.get().strip()
            if tipo == "int":
                valor = int(float(txt))
            elif tipo == "float":
                valor = float(txt)
            else:
                valor = txt
            if en_cm:
                valor = valor * cm
            setattr(cfg, attr, valor)
        # mantener coherencia de los derivados
        cfg.CUADROS_POR_HOJA = cfg.FILAS * cfg.COLUMNAS
        return cfg

    # ------------------------------------------------------- preview
    def actualizar_preview(self):
        try:
            cfg = self._construir_config()
        except ValueError as e:
            self._set_status(f"Valor invalido: {e}", error=True)
            return

        try:
            png_path = self._render_pagina_demo(cfg)
            self._cargar_imagen_preview(png_path)
            self._set_status("Vista previa actualizada.")
        except Exception as e:
            self._set_status(f"Error en vista previa: {e}", error=True)
            traceback.print_exc()

    def _render_pagina_demo(self, cfg):
        """genera un PDF temporal de 1 hoja con codigos de ejemplo y lo
        rasteriza a PNG para mostrarlo."""
        generador_obj = GeneradorEtiquetas(cfg)

        cuadros = cfg.FILAS * cfg.COLUMNAS
        # codigos de demostracion (toma de Excel real si se puede, si no, ejemplo)
        codigos_demo = self._codigos_demo(cfg, cuadros)

        tmp_pdf = os.path.join(tempfile.gettempdir(), "_preview_etiquetas.pdf")
        c = canvas.Canvas(tmp_pdf, pagesize=A4)
        ancho_hoja, alto_hoja = A4
        generador_obj._dibujar_pagina(c, codigos_demo, ancho_hoja, alto_hoja)
        c.showPage()
        c.save()

        # rasterizar primera pagina
        doc = fitz.open(tmp_pdf)
        page = doc[0]
        pix = page.get_pixmap(dpi=130)
        tmp_png = os.path.join(tempfile.gettempdir(), "_preview_etiquetas.png")
        pix.save(tmp_png)
        doc.close()
        return tmp_png

    def _codigos_demo(self, cfg, cantidad):
        """intenta leer codigos reales del Excel; si falla usa ejemplos."""
        try:
            import openpyxl
            wb = openpyxl.load_workbook(cfg.NOMBRE_EXCEL, data_only=True)
            sh = wb.active
            codigos = []
            fila = cfg.FILA_INICIAL
            while len(codigos) < cantidad:
                val = sh[f"{cfg.COLUMNA_CODIGOS}{fila}"].value
                if val is None:
                    break
                codigos.append(str(val).strip())
                fila += 1
            wb.close()
            if codigos:
                # rellenar el resto si faltan
                while len(codigos) < cantidad:
                    codigos.append("*EJEMPLO*")
                return codigos
        except Exception:
            pass
        return [f"*DEMO{i+1:03d}*" for i in range(cantidad)]

    def _cargar_imagen_preview(self, png_path):
        self._preview_pil = Image.open(png_path).copy()
        self._redibujar_preview()

    def _redibujar_preview(self):
        if not hasattr(self, "_preview_pil"):
            return
        cw = self.preview_canvas.winfo_width()
        ch = self.preview_canvas.winfo_height()
        if cw < 10 or ch < 10:
            return
        img = self._preview_pil
        escala = min(cw / img.width, ch / img.height)
        nuevo = (max(1, int(img.width * escala)), max(1, int(img.height * escala)))
        red = img.resize(nuevo, Image.LANCZOS)
        self._preview_imgtk = ImageTk.PhotoImage(red)
        self.preview_canvas.delete("all")
        self.preview_canvas.create_image(cw // 2, ch // 2, image=self._preview_imgtk)

    # ------------------------------------------------------- generar
    def generar(self):
        try:
            cfg = self._construir_config()
        except ValueError as e:
            messagebox.showerror("Valor invalido", f"Revisa los campos numericos:\n{e}")
            return

        if not os.path.exists(cfg.NOMBRE_EXCEL):
            messagebox.showerror("Excel no encontrado",
                                 f"No se encontro el archivo:\n{cfg.NOMBRE_EXCEL}")
            return

        self._set_status("Generando... (mira la consola para el detalle)")
        self.root.update_idletasks()
        try:
            generador.main(cfg)
            self._set_status("Generacion completada.")
            messagebox.showinfo(
                "Listo",
                "PDFs generados y Excel pintado correctamente.\n\n"
                f"Excel de salida: {cfg.NOMBRE_EXCEL_SALIDA}\n"
                f"Los PDFs se guardaron en:\n{self.base_dir}",
            )
        except Exception as e:
            self._set_status(f"Error: {e}", error=True)
            traceback.print_exc()
            messagebox.showerror("Error al generar", str(e))

    # ------------------------------------------------------- util
    def _set_status(self, texto, error=False):
        self.status.config(text=texto, foreground="#b00020" if error else "#1a6e1a")


# ============================================================================
#  Pestana de INGRESO MANUAL
#  No usa Excel: el usuario escribe los codigos (uno por linea), pone un
#  titulo personalizable y elige los logos y su ubicacion. Los cuadros usan
#  el mismo layout de siempre (8x3). Genera un PDF (multipagina si hace falta).
# ============================================================================

# campos personalizables de la pestana manual: (atributo, etiqueta, tipo, en_cm)
CAMPOS_MANUAL = [
    ("TITULO_HOJA", "Titulo de la hoja (arriba)", "text", False),
    ("TITULO_CUADRO", "Titulo dentro de cada cuadro", "text", False),
    ("RUTA_LOGO_UNASAM", "Logo izquierdo", "text", False),
    ("RUTA_LOGO_FACULTAD", "Logo derecho", "text", False),
    ("ALTO_IMAGENES", "Alto de logos (cm)", "float", True),
    ("MARGEN_X_LOGO_UNASAM", "Ubicacion X logo izquierdo (cm)", "float", True),
    ("MARGEN_X_LOGO_FACULTAD", "Ubicacion X logo derecho (cm)", "float", True),
    ("DISTANCIA_Y_DESDE_CODIGO", "Ubicacion Y logos (desde codigo, cm)", "float", True),
]


class PestanaManual:
    """Ingreso manual de codigos -> PDF de impresion."""

    def __init__(self, parent):
        self.root = parent
        self.base_dir = os.path.dirname(os.path.abspath(__file__))

        self.vars = {}
        self.meta = {}
        self._preview_imgtk = None

        self._construir_layout()
        self._cargar_valores_por_defecto()
        self.root.after(400, self.actualizar_preview)

    # ----------------------------------------------------------- layout
    def _construir_layout(self):
        # barra inferior
        barra = ttk.Frame(self.root)
        barra.pack(fill=tk.X, side=tk.BOTTOM)
        ttk.Separator(barra).pack(fill=tk.X)
        acciones = ttk.Frame(barra)
        acciones.pack(fill=tk.X, padx=8, pady=8)
        self.status = ttk.Label(acciones, text="Listo.", foreground="#444")
        self.status.pack(side=tk.LEFT)
        ttk.Button(acciones, text="Generar PDF",
                   command=self.generar).pack(side=tk.RIGHT, padx=4)
        ttk.Button(acciones, text="Actualizar vista previa",
                   command=self.actualizar_preview).pack(side=tk.RIGHT, padx=4)

        contenedor = ttk.PanedWindow(self.root, orient=tk.HORIZONTAL)
        contenedor.pack(fill=tk.BOTH, expand=True)

        # ----- panel izquierdo -----
        izq = ttk.Frame(contenedor, width=460)
        contenedor.add(izq, weight=0)

        # campos personalizables
        lf_cfg = ttk.LabelFrame(izq, text="Personalizacion")
        lf_cfg.pack(fill=tk.X, padx=8, pady=6)
        for attr, etiqueta, tipo, en_cm in CAMPOS_MANUAL:
            fila = ttk.Frame(lf_cfg)
            fila.pack(fill=tk.X, padx=6, pady=3)
            ttk.Label(fila, text=etiqueta, width=30, anchor="w").pack(side=tk.LEFT)
            var = tk.StringVar()
            self.vars[attr] = var
            self.meta[attr] = (tipo, en_cm)
            entry = ttk.Entry(fila, textvariable=var)
            entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
            entry.bind("<Return>", lambda e: self.actualizar_preview())
            entry.bind("<FocusOut>", lambda e: self.actualizar_preview())
            if attr in ("RUTA_LOGO_UNASAM", "RUTA_LOGO_FACULTAD"):
                ttk.Button(fila, text="...", width=3,
                           command=lambda v=var: self._buscar_logo(v)).pack(side=tk.LEFT, padx=(4, 0))

        # caja de codigos
        lf_cod = ttk.LabelFrame(izq, text="Codigos de barras (uno por linea)")
        lf_cod.pack(fill=tk.BOTH, expand=True, padx=8, pady=6)
        ttk.Label(lf_cod,
                  text="Escribe o pega un codigo por linea. Cada linea = una etiqueta.\n"
                       "El codigo se usa tal cual lo escribes.",
                  foreground="#555", justify="left").pack(anchor="w", padx=6, pady=(4, 2))

        cont_txt = ttk.Frame(lf_cod)
        cont_txt.pack(fill=tk.BOTH, expand=True, padx=6, pady=(0, 6))
        self.texto_codigos = tk.Text(cont_txt, height=12, wrap="none", font=("Consolas", 10))
        sb = ttk.Scrollbar(cont_txt, orient="vertical", command=self.texto_codigos.yview)
        self.texto_codigos.configure(yscrollcommand=sb.set)
        self.texto_codigos.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.texto_codigos.bind("<KeyRelease>", self._on_texto_cambio)

        self.lbl_conteo = ttk.Label(lf_cod, text="0 codigos - 0 hoja(s)", foreground="#1a6e1a")
        self.lbl_conteo.pack(anchor="w", padx=6, pady=(0, 4))

        # ----- panel derecho: preview -----
        der = ttk.Frame(contenedor)
        contenedor.add(der, weight=1)
        cab = ttk.Frame(der)
        cab.pack(fill=tk.X, padx=8, pady=6)
        ttk.Label(cab, text="Vista previa",
                  font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT)

        # navegacion de paginas
        self.btn_sig = ttk.Button(cab, text="Siguiente »", command=self._pagina_siguiente)
        self.btn_sig.pack(side=tk.RIGHT, padx=(4, 0))
        self.lbl_pagina = ttk.Label(cab, text="Pagina 1 de 1", width=14, anchor="center")
        self.lbl_pagina.pack(side=tk.RIGHT, padx=4)
        self.btn_ant = ttk.Button(cab, text="« Anterior", command=self._pagina_anterior)
        self.btn_ant.pack(side=tk.RIGHT, padx=(0, 4))

        self.pagina_actual = 0   # indice de pagina mostrada (0-based)
        self.total_paginas = 1

        self.preview_canvas = tk.Canvas(der, bg="#d9d9d9", highlightthickness=0)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True, padx=8, pady=(0, 6))
        self.preview_canvas.bind("<Configure>", lambda e: self._redibujar_preview())

    # ----------------------------------------------------------- helpers
    def _buscar_logo(self, var):
        ruta = filedialog.askopenfilename(
            initialdir=self.base_dir,
            filetypes=[("Imagenes", "*.png *.jpg *.jpeg"), ("Todos", "*.*")])
        if ruta:
            try:
                rel = os.path.relpath(ruta, self.base_dir)
                var.set(rel if not rel.startswith("..") else ruta)
            except ValueError:
                var.set(ruta)
            self.actualizar_preview()

    def _cargar_valores_por_defecto(self):
        base = Config()
        for attr, (tipo, en_cm) in self.meta.items():
            if attr == "TITULO_HOJA":
                valor = base.ABREVIACION_FACULTAD  # titulo de hoja por defecto
            else:
                valor = getattr(base, attr)
                if en_cm:
                    valor = round(valor / cm, 4)
            self.vars[attr].set(str(valor))
        self.actualizar_preview()

    def _leer_codigos(self):
        crudo = self.texto_codigos.get("1.0", "end")
        codigos = [ln.strip() for ln in crudo.splitlines() if ln.strip()]
        return codigos

    def _on_texto_cambio(self, _event=None):
        codigos = self._leer_codigos()
        cuadros = max(1, Config().FILAS * Config().COLUMNAS)
        hojas = (len(codigos) + cuadros - 1) // cuadros if codigos else 0
        self.lbl_conteo.config(text=f"{len(codigos)} codigos - {hojas} hoja(s)")

    def _construir_config(self):
        """Config base + personalizaciones de la pestana manual.

        El titulo de la hoja se inyecta en ABREVIACION_FACULTAD porque el
        generador dibuja ese valor como titulo superior (ver
        _dibujar_titulo_principal). No tocamos el generador original.
        """
        cfg = Config()
        for attr, var in self.vars.items():
            tipo, en_cm = self.meta[attr]
            txt = var.get().strip()
            if attr == "TITULO_HOJA":
                cfg.ABREVIACION_FACULTAD = txt
                continue
            if tipo == "float":
                valor = float(txt)
            else:
                valor = txt
            if en_cm:
                valor = valor * cm
            setattr(cfg, attr, valor)
        return cfg

    # ----------------------------------------------------------- preview
    def _pagina_anterior(self):
        if self.pagina_actual > 0:
            self.pagina_actual -= 1
            self.actualizar_preview()

    def _pagina_siguiente(self):
        if self.pagina_actual < self.total_paginas - 1:
            self.pagina_actual += 1
            self.actualizar_preview()

    def _actualizar_controles_pagina(self):
        self.lbl_pagina.config(
            text=f"Pagina {self.pagina_actual + 1} de {self.total_paginas}")
        self.btn_ant.config(
            state=tk.NORMAL if self.pagina_actual > 0 else tk.DISABLED)
        self.btn_sig.config(
            state=tk.NORMAL if self.pagina_actual < self.total_paginas - 1 else tk.DISABLED)

    def actualizar_preview(self):
        try:
            cfg = self._construir_config()
        except ValueError as e:
            self._set_status(f"Valor invalido: {e}", error=True)
            return
        try:
            codigos = self._leer_codigos()
            cuadros = cfg.FILAS * cfg.COLUMNAS
            if not codigos:
                codigos = [f"*DEMO{i+1:03d}*" for i in range(cuadros)]

            # total de paginas segun cantidad de codigos
            self.total_paginas = max(1, (len(codigos) + cuadros - 1) // cuadros)
            # mantener la pagina actual dentro de rango
            if self.pagina_actual >= self.total_paginas:
                self.pagina_actual = self.total_paginas - 1
            if self.pagina_actual < 0:
                self.pagina_actual = 0

            # codigos de la pagina que se esta mostrando
            ini = self.pagina_actual * cuadros
            codigos_pagina = codigos[ini:ini + cuadros]

            generador_obj = GeneradorEtiquetas(cfg)
            tmp_pdf = os.path.join(tempfile.gettempdir(), "_preview_manual.pdf")
            c = canvas.Canvas(tmp_pdf, pagesize=A4)
            ancho_hoja, alto_hoja = A4
            generador_obj._dibujar_pagina(c, codigos_pagina, ancho_hoja, alto_hoja)
            c.showPage()
            c.save()

            doc = fitz.open(tmp_pdf)
            pix = doc[0].get_pixmap(dpi=130)
            tmp_png = os.path.join(tempfile.gettempdir(), "_preview_manual.png")
            pix.save(tmp_png)
            doc.close()

            self._preview_pil = Image.open(tmp_png).copy()
            self._redibujar_preview()
            self._actualizar_controles_pagina()
            self._set_status("Vista previa actualizada.")
        except Exception as e:
            self._set_status(f"Error en vista previa: {e}", error=True)
            traceback.print_exc()

    def _redibujar_preview(self):
        if not hasattr(self, "_preview_pil"):
            return
        cw = self.preview_canvas.winfo_width()
        ch = self.preview_canvas.winfo_height()
        if cw < 10 or ch < 10:
            return
        img = self._preview_pil
        escala = min(cw / img.width, ch / img.height)
        nuevo = (max(1, int(img.width * escala)), max(1, int(img.height * escala)))
        red = img.resize(nuevo, Image.LANCZOS)
        self._preview_imgtk = ImageTk.PhotoImage(red)
        self.preview_canvas.delete("all")
        self.preview_canvas.create_image(cw // 2, ch // 2, image=self._preview_imgtk)

    # ----------------------------------------------------------- generar
    def generar(self):
        codigos = self._leer_codigos()
        if not codigos:
            messagebox.showwarning("Sin codigos",
                                   "Escribe al menos un codigo de barras (uno por linea).")
            return
        try:
            cfg = self._construir_config()
        except ValueError as e:
            messagebox.showerror("Valor invalido", f"Revisa los campos numericos:\n{e}")
            return

        ruta = filedialog.asksaveasfilename(
            initialdir=self.base_dir,
            defaultextension=".pdf",
            filetypes=[("PDF", "*.pdf")],
            initialfile="etiquetas_manual.pdf",
            title="Guardar PDF de etiquetas")
        if not ruta:
            return

        self._set_status("Generando PDF...")
        self.root.update_idletasks()
        try:
            generador_obj = GeneradorEtiquetas(cfg)
            c = canvas.Canvas(ruta, pagesize=A4)
            ancho_hoja, alto_hoja = A4
            cuadros = cfg.FILAS * cfg.COLUMNAS
            total_paginas = (len(codigos) + cuadros - 1) // cuadros
            for p in range(total_paginas):
                ini = p * cuadros
                pagina = codigos[ini:ini + cuadros]
                generador_obj._dibujar_pagina(c, pagina, ancho_hoja, alto_hoja)
                c.showPage()
            c.save()
            self._set_status(f"PDF generado: {os.path.basename(ruta)}")
            messagebox.showinfo(
                "Listo",
                f"PDF generado con {len(codigos)} etiquetas en {total_paginas} hoja(s).\n\n{ruta}")
        except Exception as e:
            self._set_status(f"Error: {e}", error=True)
            traceback.print_exc()
            messagebox.showerror("Error al generar", str(e))

    def _set_status(self, texto, error=False):
        self.status.config(text=texto, foreground="#b00020" if error else "#1a6e1a")


def main():
    root = tk.Tk()
    root.title("Generador de Etiquetas - UNASAM")
    root.geometry("1100x800")
    try:
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass

    notebook = ttk.Notebook(root)
    notebook.pack(fill=tk.BOTH, expand=True)

    tab_excel = ttk.Frame(notebook)
    tab_manual = ttk.Frame(notebook)
    notebook.add(tab_excel, text="  Desde Excel  ")
    notebook.add(tab_manual, text="  Ingreso manual  ")

    InterfazGenerador(tab_excel)
    PestanaManual(tab_manual)

    root.mainloop()


if __name__ == "__main__":
    main()
