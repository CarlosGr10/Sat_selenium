import tkinter as tk
from tkinter import ttk
from bot_sat import page_web

def start_process():
    start = int(inicio.get())
    finish = int(fin.get())
    the_mode = str(mode.get())

    if the_mode == 'Productivo':
        the_mode = True
    elif the_mode == 'Testeo':
        the_mode = False

    rango_personalizado = range((start - 1), finish)

    list_labels = []
    for i in rango_personalizado:
        page_web(i, the_mode)
        list_labels.append(i)
        

ventana = tk.Tk()
ventana.title("Facturacion")
ventana.config(width=400, height=300)


etiqueta_inicio = ttk.Label(text="INICIO:")
etiqueta_inicio.place(x=20, y=20)

etiqueta_fin = ttk.Label(text="FIN:")
etiqueta_fin.place(x=20, y=50)

etiqueta_modo = ttk.Label(text="ESTADO:")
etiqueta_modo.place(x=20, y=80)


inicio = ttk.Entry()
inicio.place(x=120, y=20, width=160)

fin = ttk.Entry()
fin.place(x=120, y=50, width=160)

mode = ttk.Combobox(
    values=['Testeo','Productivo']
)
mode.place(x=120, y=80, width=160)

boton_convertir = ttk.Button(text="INICIAR", command=start_process)
boton_convertir.place(x=195, y=120)

ventana.mainloop()