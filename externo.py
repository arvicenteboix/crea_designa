import tkinter as tk
import requests

def check_version(ventana, version, icon_path):

    import webbrowser
    
    try:
        url = "https://api.github.com/repos/arvicenteboix/crea_designa/releases/latest"
        response = requests.get(url, timeout=5)
        latest_release = response.json()["tag_name"]
    except:
        ventana_actualizacion = tk.Toplevel()
        ventana_actualizacion.iconbitmap(icon_path)
        ventana_actualizacion.title("Error de actualización")
        ventana_actualizacion.geometry("350x180")
        ventana_actualizacion.resizable(False, False)
        ventana_actualizacion.transient(ventana)  # La ventana de error está por encima de la principal
        ventana_actualizacion.grab_set()  # Bloquea interacción con la ventana principal hasta cerrar
        ventana_actualizacion.focus_set()

        label = tk.Label(
            ventana_actualizacion,
            text="No se pudo verificar si hay actualizaciones disponibles.\n\n"
             "Por favor, consulta la página del proyecto de vez en cuando:\n"
             "https://github.com/arvicenteboix/crea_designa/releases",
            wraplength=320,
            justify="left"
        )
        label.pack(pady=(20, 10))

        def abrir_enlace():
            webbrowser.open("https://github.com/arvicenteboix/crea_designa/releases")

        boton_enlace = tk.Button(
            ventana_actualizacion,
            text="Abrir página del proyecto",
            command=abrir_enlace,
            bg="#007bff",
            fg="white",
            font=("Arial", 10),
            relief="flat",
            padx=10,
            pady=5
        )
        boton_enlace.pack(pady=(0, 15))

        boton_cerrar = tk.Button(
            ventana_actualizacion,
            text="Cerrar",
            command=ventana_actualizacion.destroy,
            font=("Arial", 10),
            relief="flat",
            padx=10,
            pady=5
        )
        boton_cerrar.pack()

        return
    

    if latest_release != version:
        # Crear ventana personalizada con botón para abrir el enlace
        def abrir_enlace():
            webbrowser.open("https://github.com/arvicenteboix/crea_designa/releases/latest")

        ventana_actualizacion = tk.Toplevel()
        ventana_actualizacion.iconbitmap(icon_path)
        ventana_actualizacion.title("Actualización disponible")
        ventana_actualizacion.geometry("350x230")
        ventana_actualizacion.resizable(False, False)
        ventana_actualizacion.transient(ventana)  # La ventana de actualización está por encima de la principal
        ventana_actualizacion.grab_set()  # Bloquea interacción con la ventana principal hasta cerrar
        ventana_actualizacion.focus_set()
 
        label = tk.Label(
            ventana_actualizacion,
            text=f"Hay una nueva versión disponible: {latest_release}. Tienes {version}.\n\nVisita el repositorio para descargarla. Es importante que mantengas el programa actualizado para asegurar que la documentación generada cumple con las normativas vigentes.",
            wraplength=320,
            justify="left"
        )
        label.pack(pady=(20, 10))

        boton_enlace = tk.Button(
            ventana_actualizacion,
            text="Abrir página de descargas",
            command=abrir_enlace,
            bg="#007bff",
            fg="white",
            font=("Arial", 10),
            relief="flat",
            padx=10,
            pady=5
        )
        boton_enlace.pack(pady=(0, 15))

        boton_cerrar = tk.Button(
            ventana_actualizacion,
            text="Cerrar",
            command=ventana_actualizacion.destroy,
            font=("Arial", 10),
            relief="flat",
            padx=10,
            pady=5
        )
        boton_cerrar.pack()
