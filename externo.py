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
        ventana_actualizacion.title("Error d'actualització")
        ventana_actualizacion.geometry("350x180")
        ventana_actualizacion.resizable(False, False)
        ventana_actualizacion.transient(ventana)  # La ventana de error está por encima de la principal
        ventana_actualizacion.grab_set()  # Bloquea interacción con la ventana principal hasta cerrar
        ventana_actualizacion.focus_set()

        label = tk.Label(
            ventana_actualizacion,
            text="No s'ha pogut verificar si hi ha actualitzacions disponibles.\n\n"
             "Per favor, consulta la pàgina del projecte de tant en tant:\n"
             "https://github.com/arvicenteboix/crea_designa/releases",
            wraplength=320,
            justify="left"
        )
        label.pack(pady=(20, 10))

        def abrir_enlace():
            webbrowser.open("https://github.com/arvicenteboix/crea_designa/releases")

        boton_enlace = tk.Button(
            ventana_actualizacion,
            text="Obrir pàgina de descàrregues",
            width=20,
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
            text="Tancar",
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
        ventana_actualizacion.title("Actualització disponible")
        ventana_actualizacion.geometry("350x230")
        ventana_actualizacion.resizable(False, False)
        ventana_actualizacion.transient(ventana)  # La ventana de actualización está por encima de la principal
        ventana_actualizacion.grab_set()  # Bloquea interacción con la ventana principal hasta cerrar
        ventana_actualizacion.focus_set()
 
        label = tk.Label(
            ventana_actualizacion,
            text=f"Hi ha una nova versió disponible: {latest_release}. Tens {version}.\n\nVisita el repositori per descarregar-la. És important que mantingues el programa actualitzat per assegurar que la documentació generada compleix amb les normatives vigents.",
            wraplength=320,
            justify="left"
        )
        label.pack(pady=(20, 10))

        boton_enlace = tk.Button(
            ventana_actualizacion,
            text="Obrir pàgina de descàrregues",
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
