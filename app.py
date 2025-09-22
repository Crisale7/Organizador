import os
import re
import sys
from tkinter import Tk, filedialog, messagebox
from tkinter import StringVar  # status bar
from tkinter import font as tkfont
from tkinter import PhotoImage
from tkinter import TclError
from tkinter import Text
from PIL import Image, ImageTk  # <-- necesario para mostrar PNG en Labels
import tkinter.ttk as ttk

# -----------------------------
#   LÓGICA ORIGINAL (SIN CAMBIOS)
# -----------------------------

# Extensiones válidas (en minúsculas)
EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".webp"}

# Resampling compatible con varias versiones de Pillow
try:
    RESAMPLE = Image.Resampling.LANCZOS
except AttributeError:
    RESAMPLE = Image.LANCZOS

# Lista global de carpetas seleccionadas
carpetas_seleccionadas = []

# Índice de selección en el widget de listado (solo UI)
seleccion_idx = None

def seleccionar_carpeta():
    """Agregar una carpeta a la lista (selección una por una)."""
    carpeta = filedialog.askdirectory(title="Seleccionar carpeta (una a la vez)")
    if carpeta and carpeta not in carpetas_seleccionadas:
        carpetas_seleccionadas.append(carpeta)
        actualizar_listado()
        status_var.set(f"Carpeta añadida: {os.path.basename(carpeta)} • Total: {len(carpetas_seleccionadas)}")

def limpiar_lista():
    """Vaciar la lista de carpetas seleccionadas."""
    carpetas_seleccionadas.clear()
    actualizar_listado()
    status_var.set("Lista de carpetas limpiada")

def eliminar_seleccionada():
    """Eliminar solo la carpeta actualmente seleccionada en el listado (UI)."""
    global seleccion_idx
    if seleccion_idx is None:
        messagebox.showinfo("Sin selección", "Haz clic sobre una carpeta del listado para seleccionarla.")
        return
    if 0 <= seleccion_idx < len(carpetas_seleccionadas):
        borrada = os.path.basename(carpetas_seleccionadas[seleccion_idx])
        del carpetas_seleccionadas[seleccion_idx]
        seleccion_idx = None
        actualizar_listado()
        status_var.set(f"Carpeta eliminada: {borrada} • Restantes: {len(carpetas_seleccionadas)}")

def procesar_imagenes():
    """Redimensiona imagenes y guarda en la carpeta destino evitando duplicados."""
    if not carpetas_seleccionadas:
        messagebox.showerror("Error", "Debes seleccionar al menos una carpeta de origen.")
        return

    carpeta_destino = filedialog.askdirectory(title="Seleccionar carpeta de destino")
    if not carpeta_destino:
        return

    # leer tamaño (posible cambiar por el usuario)
    try:
        width = int(entry_width.get())
        height = int(entry_height.get())
    except ValueError:
        messagebox.showerror("Error", "Ingresa valores numéricos válidos para ancho y alto.")
        return

    nombres_guardados = set()  # guardaremos los nombres finales en minúscula para evitar duplicados

    total_encontradas = 0
    for carpeta in carpetas_seleccionadas:
        for root, _, files in os.walk(carpeta):
            for file in files:
                base, ext = os.path.splitext(file)
                ext_low = ext.lower()
                if ext_low in EXTENSIONS:
                    # Procesar sólo si el basename termina exactamente en "_2"
                    if re.search(r"_2$", base):
                        total_encontradas += 1
                        nuevo_base = re.sub(r"_2$", "", base)  # quita solo el sufijo final "_2"
                        nuevo_nombre = f"{nuevo_base}{ext_low}"

                        # Evitar duplicados (case-insensitive)
                        if nuevo_nombre.lower() in nombres_guardados:
                            continue

                        ruta_origen = os.path.join(root, file)
                        ruta_destino = os.path.join(carpeta_destino, nuevo_nombre)

                        try:
                            with Image.open(ruta_origen) as img:
                                # Si vamos a guardar como JPG, asegurarnos que la imagen esté en RGB
                                if ext_low in (".jpg", ".jpeg") and img.mode in ("RGBA", "LA", "P"):
                                    img = img.convert("RGB")

                                # Redimensionar con buen algoritmo
                                img = img.resize((width, height), RESAMPLE)

                                # Guardar (Pillow infiere formato por la extensión)
                                img.save(ruta_destino)
                                nombres_guardados.add(nuevo_nombre.lower())
                        except Exception as e:
                            print(f"Error procesando '{ruta_origen}': {e}")

    messagebox.showinfo(
        "Proceso finalizado",
        f"Se guardaron {len(nombres_guardados)} imágenes en:\n{carpeta_destino}\n\n"
    )
    status_var.set(f"Procesadas: {len(nombres_guardados)} • Total Encontradas: {total_encontradas}")

# -----------------------------
#   HELPERS DE UI / RECURSOS
# -----------------------------

def resource_path(relative_path: str) -> str:
    """
    Devuelve la ruta a un recurso (icono, etc.) funcionando tanto en desarrollo
    como cuando la app está congelada con PyInstaller.
    """
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)

def cargar_logo(archivo="lambol.png", alto_px=20):
    """
    Carga y escala el logo manteniendo proporción.
    Devuelve un PhotoImage listo para usar en Labels.
    """
    ruta = resource_path(archivo)
    if not os.path.exists(ruta):
        return None
    img = Image.open(ruta)
    w, h = img.size
    if h == 0:
        return None
    nuevo_ancho = int((alto_px / h) * w)
    img = img.resize((nuevo_ancho, alto_px), RESAMPLE)
    return ImageTk.PhotoImage(img)

def actualizar_listado():
    """Refresca el panel de carpetas con scroll y wrapping."""
    global seleccion_idx
    # Mantener selección si es posible
    prev = seleccion_idx

    txt_carpetas.config(state="normal")
    txt_carpetas.delete("1.0", "end")

    for i, ruta in enumerate(carpetas_seleccionadas, start=1):
        txt_carpetas.insert("end", f"{i}. {ruta}\n")

    txt_carpetas.config(state="disabled")

    # Quitar resaltado previo
    txt_carpetas.tag_delete("sel_line")
    seleccion_idx = None
    if prev is not None and 0 <= prev < len(carpetas_seleccionadas):
        # re-seleccionar misma posición si sigue existiendo
        select_line(prev)

def on_click_listado(event):
    """Detecta línea clicada en el Text y la resalta como selección."""
    index = txt_carpetas.index(f"@{event.x},{event.y}")
    line = int(index.split(".")[0])  # 1-based
    # Mapear línea a índice de lista
    idx = line - 1
    if 0 <= idx < len(carpetas_seleccionadas):
        select_line(idx)
        status_var.set(f"Seleccionado: {carpetas_seleccionadas[idx]}")

def select_line(idx: int):
    """Resalta visualmente la línea 'idx' (0-based) en el Text."""
    global seleccion_idx
    seleccion_idx = idx
    txt_carpetas.tag_delete("sel_line")
    start = f"{idx+1}.0 linestart"
    end = f"{idx+1}.0 lineend"
    txt_carpetas.tag_add("sel_line", start, end)
    txt_carpetas.tag_config("sel_line", background="#e0ebff")

# -----------------------------
#   UI MEJORADA (ESTÉTICA)
# -----------------------------

root = Tk()
root.title("Procesador de imágenes  •  Organizador")
root.geometry("780x560")
root.minsize(720, 520)

# Tema y estilos
style = ttk.Style()
for candidate in ("vista", "xpnative", "clam"):
    try:
        style.theme_use(candidate)
        break
    except TclError:
        continue

PRIMARY = "#2563eb"
BG_CARD = "#ffffff"
BG_APP = "#f5f6fb"
TXT_MUTED = "#636a77"

try:
    root.configure(bg=BG_APP)
except Exception:
    pass

# Tipografías
title_font = tkfont.Font(size=13, weight="bold")
subtitle_font = tkfont.Font(size=10)
label_font = tkfont.Font(size=10)
mono_font = tkfont.Font(family="Consolas", size=9)

# Estilos ttk
style.configure("Card.TFrame", background=BG_CARD, relief="flat")
style.configure("App.TFrame", background=BG_APP)
style.configure("Header.TLabel", background=BG_APP, font=title_font, foreground="#0f172a")
style.configure("SubHeader.TLabel", background=BG_APP, font=subtitle_font, foreground=TXT_MUTED)
style.configure("TLabel", background=BG_CARD, font=label_font)
style.configure("Muted.TLabel", background=BG_CARD, font=label_font, foreground=TXT_MUTED)
style.configure("TButton", padding=8)
style.map("TButton", background=[("active", PRIMARY)], foreground=[("active", "white")])

# Contenedor principal
app = ttk.Frame(root, style="App.TFrame", padding=16)
app.pack(fill="both", expand=True)

# Header (solo textos; sin logo)
header = ttk.Frame(app, style="App.TFrame")
header.pack(fill="x", pady=(0, 10))
ttk.Label(header, text="Procesador de imágenes", style="Header.TLabel").pack(anchor="w")
ttk.Label(header, text="Selecciona carpetas, define tamaño y procesa imágenes.",
          style="SubHeader.TLabel").pack(anchor="w", pady=(2, 4))

# Grid
grid = ttk.Frame(app, style="App.TFrame")
grid.pack(fill="both", expand=True)
grid.columnconfigure(0, weight=3, uniform="cols")
grid.columnconfigure(1, weight=2, uniform="cols")
grid.rowconfigure(0, weight=1)

# ---- Card: Orígenes ----
card_origen = ttk.Frame(grid, style="Card.TFrame", padding=16)
card_origen.grid(row=0, column=0, sticky="nsew", padx=(0, 10))
card_origen.rowconfigure(3, weight=1)

ttk.Label(card_origen, text="Carpetas de origen").grid(row=0, column=0, columnspan=3, sticky="w")
ttk.Separator(card_origen, orient="horizontal").grid(row=1, column=0, columnspan=3, sticky="ew", pady=6)

btns_origen = ttk.Frame(card_origen, style="Card.TFrame")
btns_origen.grid(row=2, column=0, columnspan=3, sticky="w", pady=(2, 8))
ttk.Button(btns_origen, text="➕ Agregar carpeta", command=seleccionar_carpeta).grid(row=0, column=0, padx=(0, 8))
ttk.Button(btns_origen, text="🗑️ Limpiar lista", command=limpiar_lista).grid(row=0, column=1, padx=(0, 8))
ttk.Button(btns_origen, text="❌ Eliminar seleccionada", command=eliminar_seleccionada).grid(row=0, column=2)

# Panel con Text + Scrollbar (wrap palabras)
panel = ttk.Frame(card_origen, style="Card.TFrame")
panel.grid(row=3, column=0, columnspan=3, sticky="nsew")

scrollbar = ttk.Scrollbar(panel, orient="vertical")
scrollbar.pack(side="right", fill="y")

txt_carpetas = Text(
    panel, wrap="word", height=12, relief="solid", bd=1,
    background="#ffffff", foreground="#0f172a", font=mono_font
)
txt_carpetas.pack(fill="both", expand=True, padx=2, pady=2)
txt_carpetas.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=txt_carpetas.yview)

# Click para seleccionar línea
txt_carpetas.bind("<Button-1>", on_click_listado)

# ---- Card: Tamaño & Acción ----
card_size = ttk.Frame(grid, style="Card.TFrame", padding=16)
card_size.grid(row=0, column=1, sticky="nsew")

ttk.Label(card_size, text="Configuración de salida").grid(row=0, column=0, columnspan=2, sticky="w")
ttk.Separator(card_size, orient="horizontal").grid(row=1, column=0, columnspan=2, sticky="ew", pady=6)

ttk.Label(card_size, text="Tamaño de salida (px):", style="Muted.TLabel").grid(row=2, column=0, columnspan=2, sticky="w", pady=(0, 6))

ttk.Label(card_size, text="Ancho:").grid(row=3, column=0, sticky="w", padx=(0, 8))
entry_width = ttk.Entry(card_size, width=10)
entry_width.insert(0, "720")
entry_width.grid(row=3, column=1, sticky="ew")

ttk.Label(card_size, text="Alto:").grid(row=4, column=0, sticky="w", padx=(0, 8), pady=(6, 0))
entry_height = ttk.Entry(card_size, width=10)
entry_height.insert(0, "500")
entry_height.grid(row=4, column=1, sticky="ew", pady=(6, 0))

ttk.Label(card_size, text="", style="Card.TFrame").grid(row=5, column=0, columnspan=2, pady=8)

btn_procesar = ttk.Button(card_size, text="🚀 Procesar imágenes", command=procesar_imagenes)
btn_procesar.grid(row=6, column=0, columnspan=2, sticky="ew", pady=(4, 0))

card_size.columnconfigure(1, weight=1)

# Footer / status bar (CON LOGO)
status_var = StringVar(value="Listo")
status = ttk.Frame(app, style="App.TFrame")
status.pack(fill="x", pady=(10, 0))
ttk.Separator(status, orient="horizontal").pack(fill="x", pady=(0, 6))

# Cargar logo para el footer (más pequeño, p. ej. 20 px alto)
root.logo_status = cargar_logo("lambol.png", alto_px=20)

# Etiqueta con imagen + texto (compound='left' muestra el icono antes del texto)
lbl_status = ttk.Label(status, textvariable=status_var, style="SubHeader.TLabel")
if root.logo_status:
    lbl_status.configure(image=root.logo_status, compound="left", padding=(6, 0))
lbl_status.pack(anchor="w")

# Icono de la app (NO usamos lambol.png aquí; solo icono.ico/icono.png si existen)
try:
    ico_path = resource_path("icono.ico")
    if os.path.exists(ico_path):
        root.iconbitmap(ico_path)
    else:
        png_path = resource_path("icono.png")
        if os.path.exists(png_path):
            png_icon = PhotoImage(file=png_path)
            root.iconphoto(True, png_icon)
except Exception:
    pass

# Inicializar listado vacío
def _init_after_layout():
    actualizar_listado()

root.after(0, _init_after_layout)
root.mainloop()
