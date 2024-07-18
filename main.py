import tkinter as tk
from tkinter import messagebox
from openpyxl import *
import re
import os

nombre_archivo = 'Datos.xlsx'
if os.path.exists(nombre_archivo):
    wb = load_workbook(nombre_archivo)
    ws = wb.active
else:
    #create a file is not exist
    wb = Workbook()
    ws = wb.active
    ws.append(["Nombre", "Edad", "Email", "Telefono", "Dirección"]) #Agregamos una fila


def guardar_Datos():
    nombre = entry_nombre.get()
    edad = entry_edad.get()
    email = entry_mail.get()
    telefono = entry_telefono.get()
    direccion = entry_direccion.get()

    if not nombre or not edad or not edad or not email or not telefono or not direccion:
        messagebox.showwarning("Adventencia", "Todos los campos, deben ser diligenciados")
        return
    try:
        edad = int(edad)
        telefono = int(telefono)
    except ValueError:
         messagebox.showwarning("Adventencia", "Los campos telefono y edad son numericos")
         return
    if not re.match(r"[^@]+@[^@]+\.[^@]+", email):
        messagebox.showwarning("Advertencia", "Correo no valido")
        return
    
    ws.append([nombre, edad, email, telefono, direccion])
    wb.save(nombre_archivo)
    messagebox.showinfo("Informacion", "Datos guardados con exito")

    entry_nombre.delete(0,tk.END)
    entry_edad.delete(0,tk.END)
    entry_mail.delete(0,tk.END)
    entry_telefono.delete(0,tk.END)
    entry_direccion.delete(0,tk.END)
#wb.save('datos.xlsx') #Guardamos el archivo

root = tk.Tk()
root.title("Formulario de acceso de datos")
root.configure(bg='#4B6587')
label_style = {"bg": "#4B6587", "fg": "white"}
entry_style = {"bg": "#D3D3D3", "fg": "black"}

label_nombre = tk.Label(root, text="Nombre", **label_style) #Stilo de la etiqute nombre
label_nombre.grid(row=0, column=0, padx=10, pady=5)
entry_nombre = tk.Entry(root, **entry_style)
entry_nombre.grid(row=0, column=1, padx=10, pady=5)

label_edad = tk.Label(root, text="Edad", **label_style) #Stilo de la etiqute nombre
label_edad.grid(row=1, column=0, padx=10, pady=5)
entry_edad = tk.Entry(root, **entry_style)
entry_edad.grid(row=1, column=1, padx=10, pady=5)

label_mail = tk.Label(root, text="Email", **label_style) #Stilo de la etiqute nombre
label_mail.grid(row=2, column=0, padx=10, pady=5)
entry_mail = tk.Entry(root, **entry_style)
entry_mail.grid(row=2, column=1, padx=10, pady=5)

label_telefono = tk.Label(root, text="Telefono", **label_style) #Stilo de la etiqute nombre
label_telefono.grid(row=3, column=0, padx=10, pady=5)
entry_telefono = tk.Entry(root, **entry_style)
entry_telefono.grid(row=3, column=1, padx=10, pady=5)

label_direccion = tk.Label(root, text="Dirección", **label_style) #Stilo de la etiqute nombre
label_direccion.grid(row=4, column=0, padx=10, pady=5)
entry_direccion = tk.Entry(root, **entry_style)
entry_direccion.grid(row=4, column=1, padx=10, pady=5)

boton_guardar = tk.Button(root, text="Guardar", command =guardar_Datos, 
                          bg="#6D8299", fg='white', width=20)
boton_guardar.grid(row=5, column=0, columnspan=2, padx=10, pady=10)

root.mainloop()