# main.py
"""
Punto de entrada simplificado para ContaFlow usando tkinter nativo.
Inicializa y ejecuta la ventana principal sin dependencias externas.
"""
# Archivos relacionados: main_window.py

import tkinter as tk
from tkinter import messagebox
import signal
import sys
import os
import threading


def signal_handler(signum, frame):
    """Maneja las señales de interrupción para cerrar correctamente."""
    print("\n🔄 Cerrando ContaFlow...")
    try:
        sys.exit(0)
    except:
        os._exit(0)


def handle_thread_exception(args):
    """Maneja excepciones no capturadas en threads."""
    print(f"❌ Error en thread: {args.exc_type.__name__}: {args.exc_value}")


def main():
    """Función principal simplificada que inicializa ContaFlow."""
    print("🚀 Iniciando ContaFlow...")

    try:
        # Configurar manejo de señales
        signal.signal(signal.SIGINT, signal_handler)
        signal.signal(signal.SIGTERM, signal_handler)

        # Configurar manejo de excepciones en threads
        threading.excepthook = handle_thread_exception

        print("📱 Creando ventana principal...")

        # Importar después de configurar para evitar problemas
        from ui.main_window import MainWindow

        # Crear y configurar la ventana principal
        app = MainWindow()

        print("✅ ContaFlow iniciado correctamente")
        print("🎨 Usando tema nativo del sistema")

        # Iniciar el bucle principal
        app.mainloop()

    except KeyboardInterrupt:
        print("\n⚠️ Aplicación interrumpida por el usuario")
        sys.exit(0)

    except ImportError as import_error:
        print(f"❌ Error de importación: {import_error}")
        print("💡 Verifique que todos los archivos estén presentes")

        # Mostrar error en ventana si es posible
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Error de ContaFlow",
                                 f"Error de importación:\n{import_error}\n\nVerifique que todos los archivos estén presentes.")
            root.destroy()
        except:
            pass

        sys.exit(1)

    except Exception as e:
        print(f"❌ Error crítico en la aplicación: {e}")
        print(f"   Tipo de error: {type(e).__name__}")

        # Mostrar error en ventana
        try:
            root = tk.Tk()
            root.withdraw()
            messagebox.showerror("Error de ContaFlow",
                                 f"Error crítico:\n{type(e).__name__}: {e}\n\nLa aplicación se cerrará.")
            root.destroy()
        except:
            pass

        sys.exit(1)


if __name__ == "__main__":
    main()