import os
import re
from tkinter import Tk, filedialog, Label, Button, Entry, messagebox, Frame
from PIL import Image

# Extensiones válidas (en minúsculas)
EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif", ".tiff", ".webp"}

# Resampling compatible con varias versiones de Pillow
try:
    RESAMPLE = Image.Resampling.LANCZOS
except AttributeError:
    RESAMPLE = Image.LANCZOS

# Lista global de carpetas seleccionadas
carpetas_seleccionadas = []

def seleccionar_carpeta():
    """Agregar una carpeta a la lista (selección una por una)."""
    carpeta = filedialog.askdirectory(title="Seleccionar carpeta (una a la vez)")
    if carpeta and carpeta not in carpetas_seleccionadas:
        carpetas_seleccionadas.append(carpeta)
        lbl_carpetas.config(text="\n".join(carpetas_seleccionadas))

def limpiar_lista():
    """Vaciar la lista de carpetas seleccionadas."""
    carpetas_seleccionadas.clear()
    lbl_carpetas.config(text="")

def procesar_imagenes():
    """Busca archivos que terminan en _2 (antes de la extensión), los renombra quitando solo ese sufijo,
       redimensiona y guarda en la carpeta destino evitando duplicados."""
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

    for carpeta in carpetas_seleccionadas:
        for root, _, files in os.walk(carpeta):
            for file in files:
                base, ext = os.path.splitext(file)
                ext_low = ext.lower()
                if ext_low in EXTENSIONS:
                    # Procesar sólo si el basename termina exactamente en "_2"
                    if re.search(r"_2$", base):
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

    messagebox.showinfo("Proceso finalizado", f"Se guardaron {len(nombres_guardados)} imágenes en:\n{carpeta_destino}")

# --- Interfaz sencilla con Tkinter ---
root = Tk()
root.title("Procesador de imágenes (_2 → quitar sufijo)")
root.geometry("600x420")

Label(root, text="Selecciona carpetas de origen (puedes agregar varias):").pack(pady=6)
frame1 = Frame(root)
frame1.pack()
Button(frame1, text="Agregar carpeta", command=seleccionar_carpeta).grid(row=0, column=0, padx=6)
Button(frame1, text="Limpiar lista", command=limpiar_lista).grid(row=0, column=1, padx=6)

lbl_carpetas = Label(root, text="", justify="left", anchor="w")
lbl_carpetas.pack(fill="both", padx=10, pady=6)

Label(root, text="Tamaño de salida (px):").pack(pady=(8,0))
size_frame = Frame(root)
size_frame.pack(pady=4)
Label(size_frame, text="Ancho:").grid(row=0, column=0, padx=4)
entry_width = Entry(size_frame, width=8); entry_width.insert(0, "720"); entry_width.grid(row=0, column=1)
Label(size_frame, text="Alto:").grid(row=0, column=2, padx=8)
entry_height = Entry(size_frame, width=8); entry_height.insert(0, "500"); entry_height.grid(row=0, column=3)

Button(root, text="Procesar imágenes", command=procesar_imagenes, width=20, height=2).pack(pady=18)

root.mainloop()
