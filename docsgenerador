import tkinter as tk
from tkinter import messagebox
from docx import Document
import shutil
import os
import sys

# Obtener la ruta del recurso
def ruta_recurso(rel_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, rel_path)
    return os.path.join(os.path.abspath("."), rel_path)

# Generar y abrir el documento
def generar_documento_desde_plantilla(nombre_plantilla, nombre_salida):
    try:
        plantilla_path = ruta_recurso(nombre_plantilla)
        shutil.copy(plantilla_path, nombre_salida)
        doc = Document(nombre_salida)
        doc.save(nombre_salida)
        os.startfile(nombre_salida)
    except Exception as e:
        messagebox.showerror("Error", f"No se pudo generar el documento.\n{str(e)}")

# Escoger tipo de documento vertical
def seleccionar_vertical():
    def generar():
        tipo = lista_tipos.get(lista_tipos.curselection())
        plantilla_map = {
            "Memo": "memo.docx",
            "Solicitud": "solicitud.docx",
            "Orden de salida": "orden_de_salida.docx",
            "Informe técnico": "informe_tecnico.docx"
        }
        plantilla = plantilla_map[tipo]
        nombre_salida = f"{tipo.replace(' ', '_').lower()}_vertical.docx"
        generar_documento_desde_plantilla(plantilla, nombre_salida)

    ventana_vertical = tk.Toplevel()
    ventana_vertical.title("Selecciona tipo de documento vertical")
    ventana_vertical.geometry("300x250")

    etiqueta = tk.Label(ventana_vertical, text="Tipo de documento:", font=("Agency FB", 14))
    etiqueta.pack(pady=10)

    tipos = ["Memo", "Solicitud", "Orden de salida", "Informe técnico"]
    lista_tipos = tk.Listbox(ventana_vertical, height=4, selectmode=tk.SINGLE, font=("Arial", 12))
    for tipo in tipos:
        lista_tipos.insert(tk.END, tipo)
    lista_tipos.pack(pady=10)

    btn_generar = tk.Button(ventana_vertical, text="Generar", command=generar, width=20)
    btn_generar.pack(pady=10)

# Manejo de la selección horizontal
def seleccionar_horizontal():
    generar_documento_desde_plantilla("Plantilla_H.docx", "documento_horizontal.docx")

# Ventana principal
ventana = tk.Tk()
ventana.title("Generador de Words")
ventana.geometry("300x200")

etiqueta = tk.Label(ventana, text="Seleccione orientación:", font=("Agency FB", 14))
etiqueta.pack(pady=10)

btn_vertical = tk.Button(ventana, text="Vertical", command=seleccionar_vertical, width=20)
btn_vertical.pack(pady=5)

btn_horizontal = tk.Button(ventana, text="Horizontal", command=seleccionar_horizontal, width=20)
btn_horizontal.pack(pady=5)

ventana.mainloop()
