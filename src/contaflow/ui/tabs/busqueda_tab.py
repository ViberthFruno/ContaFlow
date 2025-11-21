# busqueda_tab.py
"""
Sub-pestaña simplificada para configuración de búsqueda de correos 'Cargador' con archivos Excel.
Configuración fija optimizada para el nuevo sistema robusto de procesamiento.
Diseño moderno optimizado.
"""
# Archivos relacionados: config_manager.py, theme_manager.py

import tkinter as tk
from tkinter import ttk, filedialog
import os
from config.config_manager import ConfigManager
from ui.theme_manager import ModernTheme


class BusquedaTab:
    """Sub-pestaña simplificada para configuración de búsqueda de correos 'Cargador'."""

    def __init__(self, parent):
        """Inicializa la sub-pestaña con configuración fija para 'Cargador'."""
        self.parent = parent
        self.is_visible = False
        self.config_manager = ConfigManager()

        # Variable para la carpeta de descarga
        self.download_folder_var = tk.StringVar()

        # Referencia al widget de estado
        self.status_label = None

        # Crear la interfaz de usuario
        self.create_interface()
        self.load_search_config()

    def create_interface(self):
        """Crea la interfaz de usuario moderna optimizada."""
        # Frame principal que contendrá todos los elementos
        main_frame = ttk.Frame(self.parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Frame para la configuración requerida con estilo moderno
        config_frame = ttk.LabelFrame(main_frame, text="📁 Configuración Requerida",
                                     padding=15, style="Modern.TLabelframe")
        config_frame.pack(fill=tk.BOTH, expand=True)

        # --- Sección de Carpeta de Descarga ---
        self.create_download_folder_section(config_frame)

        # --- Sección de Estado de Configuración ---
        status_card = tk.Frame(config_frame, bg=ModernTheme.BG_SURFACE,
                              highlightbackground=ModernTheme.BORDER_LIGHT,
                              highlightthickness=1)
        status_card.pack(fill=tk.X, pady=(20, 10))

        tk.Label(status_card, text="Estado de Configuración",
                font=ModernTheme.FONT_SUBHEADING,
                bg=ModernTheme.BG_SURFACE,
                fg=ModernTheme.TEXT_PRIMARY).pack(pady=(8, 5))

        self.status_label = tk.Label(status_card, text="🔴 Carpeta no configurada",
                                     font=ModernTheme.FONT_NORMAL,
                                     bg=ModernTheme.BG_SURFACE,
                                     fg=ModernTheme.TEXT_PRIMARY,
                                     wraplength=350)
        self.status_label.pack(fill=tk.X, pady=(0, 8))

        # --- Sección de Botones de Control ---
        buttons_frame = ttk.Frame(config_frame)
        buttons_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(10, 0))
        buttons_frame.columnconfigure((0, 1, 2), weight=1)

        save_btn = ttk.Button(buttons_frame, text="💾 Guardar",
                              command=self.save_search_config,
                              style="Success.TButton")
        save_btn.grid(row=0, column=0, sticky=tk.EW, padx=(0, 5), ipady=5)

        reset_btn = ttk.Button(buttons_frame, text="🔄 Restaurar",
                               command=self.set_default_folder,
                               style="TButton")
        reset_btn.grid(row=0, column=1, sticky=tk.EW, padx=5, ipady=5)

        clear_btn = ttk.Button(buttons_frame, text="🗑️ Limpiar",
                               command=self.clear_search_config,
                               style="TButton")
        clear_btn.grid(row=0, column=2, sticky=tk.EW, padx=(5, 0), ipady=5)


    def create_download_folder_section(self, parent):
        """Crea la sección moderna para configurar la carpeta de descarga."""
        download_frame = ttk.Frame(parent)
        download_frame.pack(fill=tk.X, anchor="n")

        download_label = ttk.Label(download_frame, text="Carpeta de Descarga:",
                                   font=ModernTheme.FONT_SUBHEADING)
        download_label.pack(anchor=tk.W, pady=(0, 8))

        # Frame para el campo de entrada y el botón de búsqueda
        entry_frame = ttk.Frame(download_frame)
        entry_frame.pack(fill=tk.X)
        entry_frame.grid_columnconfigure(0, weight=1)

        download_entry = ttk.Entry(entry_frame, textvariable=self.download_folder_var,
                                   font=ModernTheme.FONT_NORMAL)
        download_entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))

        browse_btn = ttk.Button(entry_frame, text="📁", width=4,
                                command=self.browse_folder,
                                style="TButton")
        browse_btn.grid(row=0, column=1)

    def browse_folder(self):
        """Abre un diálogo para seleccionar una carpeta."""
        folder_path = filedialog.askdirectory(
            title="Seleccionar carpeta de descarga",
            initialdir=os.path.expanduser("~")
        )
        if folder_path:
            self.download_folder_var.set(folder_path)
            self.update_status("🟡 Carpeta seleccionada. Guarde la configuración.", "orange")

    def set_default_folder(self):
        """Establece la carpeta de descarga por defecto."""
        default_folder = os.path.join(os.path.expanduser("~"), "Downloads", "ContaFlow_Cargador")
        self.download_folder_var.set(default_folder)
        self.update_status("🟡 Carpeta por defecto establecida. Guarde la configuración.", "orange")

    def save_search_config(self):
        """Guarda la configuración de la carpeta de descarga."""
        download_folder = self.download_folder_var.get().strip()

        if not download_folder:
            self.update_status("🔴 Debe especificar una carpeta de descarga.", "red")
            return

        try:
            os.makedirs(download_folder, exist_ok=True)
            config_data = {
                "subject": "Cargador",
                "search_type": "Contiene",
                "today_only": True,
                "attachments_only": True,
                "excel_files": True,
                "download_folder": download_folder
            }
            self.config_manager.update_config({"search_criteria": config_data})
            self.update_status("🟢 Configuración guardada correctamente.", "green")
        except Exception as e:
            self.update_status(f"🔴 Error al guardar: {str(e)}", "red")

    def load_search_config(self):
        """Carga la configuración existente o establece los valores por defecto."""
        try:
            config = self.config_manager.load_config()
            search_criteria = config.get("search_criteria", {})
            folder = search_criteria.get("download_folder", "")

            if folder:
                self.download_folder_var.set(folder)
                self.update_status("🟢 Configuración cargada.", "green")
            else:
                self.set_default_folder()
        except Exception as e:
            print(f"Error cargando configuración de búsqueda: {e}")
            self.set_default_folder()

    def clear_search_config(self):
        """Limpia el campo de la carpeta de descarga."""
        self.download_folder_var.set("")
        self.update_status("🔴 Carpeta no configurada.", "red")

    def update_status(self, message, color):
        """Actualiza el mensaje y color del texto de estado."""
        if self.status_label:
            self.status_label.config(text=message, foreground=color)

    def validate_config(self):
        """Valida si la configuración actual es válida."""
        folder = self.download_folder_var.get().strip()
        if not folder:
            return False, "Debe configurar una carpeta de descarga."
        if not os.path.isdir(os.path.dirname(folder)):
            return False, "La ruta de la carpeta base no es válida."
        return True, "Configuración válida."

    # --- Métodos adicionales (no modificados) ---
    def get_current_config(self):
        return {
            "subject": "Cargador", "search_type": "Contiene", "today_only": True,
            "attachments_only": True, "excel_files": True,
            "download_folder": self.download_folder_var.get().strip()
        }

    def get_config_summary(self):
        is_valid, message = self.validate_config()
        config = self.get_current_config()
        return {
            "valid": is_valid, "message": message, "subject": config["subject"],
            "download_folder": config["download_folder"], "system_type": "Configuración Fija Optimizada",
            "monitoring_interval": "1 minuto (fijo)", "search_robustness": "Búsqueda sin estado UNSEEN"
        }

    def show(self):
        if not self.is_visible:
            self.is_visible = True

    def hide(self):
        if self.is_visible:
            self.is_visible = False