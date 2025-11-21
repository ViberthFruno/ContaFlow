# email_config_modals.py
"""
Modales para configuración de email, destinatarios, búsqueda, XML y exclusiones de combustible.
Usados desde la pestaña de Automatización.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
import re
import os
import unicodedata
from typing import List
from email_manager import EmailManager
from config_manager import ConfigManager
from theme_manager import ModernTheme


class EmailConfigModal:
    """Modal para configuración de credenciales de email."""

    def __init__(self, parent):
        """Inicializa el modal de configuración de email."""
        self.parent = parent
        self.email_manager = EmailManager()
        self.config_manager = ConfigManager()

        # Variables de credenciales
        self.provider_var = tk.StringVar(value="Gmail")
        self.email_var = tk.StringVar()
        self.password_var = tk.StringVar()

        # Crear ventana modal
        self.window = tk.Toplevel(parent)
        self.window.title("⚙️ Configuración de Email")
        self.window.geometry("500x450")
        self.window.resizable(False, False)

        # Hacer modal (bloquear ventana principal)
        self.window.transient(parent)
        self.window.grab_set()

        # Centrar ventana
        self._center_window()

        # Crear interfaz
        self.create_interface()

        # Cargar configuración existente
        self.load_existing_config()

    def _center_window(self):
        """Centra la ventana modal en la pantalla."""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")

    def create_interface(self):
        """Crea la interfaz del modal."""
        # Frame principal con padding
        main_frame = ttk.Frame(self.window, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Título
        title_label = tk.Label(main_frame, text="📧 Configuración de Email",
                              font=ModernTheme.FONT_HEADING,
                              fg=ModernTheme.PRIMARY)
        title_label.pack(pady=(0, 20))

        # Frame para campos
        fields_frame = ttk.Frame(main_frame)
        fields_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        # Proveedor
        ttk.Label(fields_frame, text="Proveedor de Correo:",
                 font=ModernTheme.FONT_NORMAL).pack(anchor=tk.W, pady=(0, 5))
        provider_combo = ttk.Combobox(fields_frame, textvariable=self.provider_var,
                                      values=["Gmail", "Outlook", "Yahoo", "Otro"],
                                      state="readonly", font=ModernTheme.FONT_NORMAL)
        provider_combo.pack(fill=tk.X, pady=(0, 15))

        # Email
        ttk.Label(fields_frame, text="Correo Electrónico:",
                 font=ModernTheme.FONT_NORMAL).pack(anchor=tk.W, pady=(0, 5))
        email_entry = ttk.Entry(fields_frame, textvariable=self.email_var,
                               font=ModernTheme.FONT_NORMAL)
        email_entry.pack(fill=tk.X, pady=(0, 15))

        # Contraseña
        ttk.Label(fields_frame, text="Contraseña:",
                 font=ModernTheme.FONT_NORMAL).pack(anchor=tk.W, pady=(0, 5))
        password_entry = ttk.Entry(fields_frame, textvariable=self.password_var,
                                   show="*", font=ModernTheme.FONT_NORMAL)
        password_entry.pack(fill=tk.X, pady=(0, 15))

        # Nota informativa
        note_frame = tk.Frame(fields_frame, bg=ModernTheme.INFO,
                             highlightbackground=ModernTheme.SECONDARY,
                             highlightthickness=1)
        note_frame.pack(fill=tk.X, pady=(10, 0))

        note_label = tk.Label(note_frame,
                             text="💡 Para Gmail usa una contraseña de aplicación",
                             fg=ModernTheme.TEXT_WHITE, bg=ModernTheme.INFO,
                             font=ModernTheme.FONT_SMALL, justify=tk.CENTER, pady=8)
        note_label.pack()

        # Estado
        self.status_label = tk.Label(fields_frame, text="",
                                     font=ModernTheme.FONT_SMALL,
                                     fg=ModernTheme.TEXT_SECONDARY)
        self.status_label.pack(pady=(10, 0))

        # Frame de botones
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, side=tk.BOTTOM)

        # Botón probar conexión
        test_btn = ttk.Button(buttons_frame, text="🔍 Probar Conexión",
                             command=self.test_connection,
                             style="TButton")
        test_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5), ipady=8)

        # Botón guardar
        save_btn = ttk.Button(buttons_frame, text="💾 Guardar",
                             command=self.save_config,
                             style="Success.TButton")
        save_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5, ipady=8)

        # Botón cancelar
        cancel_btn = ttk.Button(buttons_frame, text="❌ Cancelar",
                               command=self.window.destroy,
                               style="TButton")
        cancel_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 0), ipady=8)

    def load_existing_config(self):
        """Carga configuración existente."""
        try:
            config = self.config_manager.load_config()
            if config:
                self.provider_var.set(config.get("provider", "Gmail"))
                self.email_var.set(config.get("email", ""))
                self.password_var.set(config.get("password", ""))

                if config.get("email"):
                    self.update_status("🟡 Configuración cargada", "orange")
        except Exception as e:
            print(f"Error cargando configuración: {e}")

    def test_connection(self):
        """Prueba la conexión de email."""
        credentials = self._get_credentials_data()
        if not all(credentials.values()):
            return self.update_status("🔴 Complete todos los campos", "red")

        self.update_status("🔄 Probando conexión...", "orange")

        def test_thread():
            try:
                success, message = self.email_manager.test_connection(
                    credentials["provider"], credentials["email"], credentials["password"]
                )

                def update_ui():
                    color = "green" if success else "red"
                    icon = "🟢" if success else "🔴"
                    self.update_status(f"{icon} {message}", color)

                self.window.after(0, update_ui)

            except Exception as e:
                self.window.after(0, lambda: self.update_status(f"🔴 Error: {str(e)}", "red"))

        threading.Thread(target=test_thread, daemon=True).start()

    def save_config(self):
        """Guarda la configuración de email."""
        credentials = self._get_credentials_data()
        if not all(credentials.values()):
            return self.update_status("🔴 Complete todos los campos", "red")

        # Validar formato de email
        if not self._validate_email_format(credentials["email"]):
            return self.update_status("🔴 Formato de email inválido", "red")

        try:
            # Cargar configuración existente para no sobreescribir otros datos
            existing_config = self.config_manager.load_config() or {}

            # Actualizar solo las credenciales
            existing_config.update(credentials)

            # Guardar
            self.config_manager.save_config(existing_config)

            self.update_status("🟢 Configuración guardada", "green")
            messagebox.showinfo("Éxito", "Configuración de email guardada correctamente")

        except Exception as e:
            self.update_status(f"🔴 Error: {str(e)}", "red")
            messagebox.showerror("Error", f"Error al guardar configuración: {str(e)}")

    def _get_credentials_data(self):
        """Obtiene los datos de credenciales actuales."""
        return {
            "provider": self.provider_var.get(),
            "email": self.email_var.get().strip(),
            "password": self.password_var.get().strip()
        }

    def _validate_email_format(self, email):
        """Valida formato de email."""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, email.strip()))

    def update_status(self, message, color):
        """Actualiza el estado con colores modernos."""
        if self.status_label:
            color_map = {
                'green': ModernTheme.SUCCESS,
                'red': ModernTheme.DANGER,
                'orange': ModernTheme.WARNING,
                'blue': ModernTheme.INFO
            }
            modern_color = color_map.get(color, color)
            self.status_label.config(text=message, fg=modern_color)


class RecipientsConfigModal:
    """Modal para configuración de destinatarios y CC."""

    def __init__(self, parent):
        """Inicializa el modal de configuración de destinatarios."""
        self.parent = parent
        self.config_manager = ConfigManager()

        # Variables de destinatarios
        self.main_email_var = tk.StringVar()
        self.cc_text = None  # Widget Text para CC separados por comas

        # Crear ventana modal
        self.window = tk.Toplevel(parent)
        self.window.title("📧 Configurar Destinatarios")
        self.window.geometry("550x450")
        self.window.resizable(False, False)

        # Hacer modal (bloquear ventana principal)
        self.window.transient(parent)
        self.window.grab_set()

        # Centrar ventana
        self._center_window()

        # Crear interfaz
        self.create_interface()

        # Cargar configuración existente
        self.load_existing_config()

    def _center_window(self):
        """Centra la ventana modal en la pantalla."""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")

    def create_interface(self):
        """Crea la interfaz del modal."""
        # Frame principal con padding
        main_frame = ttk.Frame(self.window, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Título
        title_label = tk.Label(main_frame, text="👥 Configurar Destinatarios",
                              font=ModernTheme.FONT_HEADING,
                              fg=ModernTheme.PRIMARY)
        title_label.pack(pady=(0, 20))

        # Frame para campos
        fields_frame = ttk.Frame(main_frame)
        fields_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        # Destinatario principal
        ttk.Label(fields_frame, text="📧 Destinatario Principal:",
                 font=ModernTheme.FONT_NORMAL).pack(anchor=tk.W, pady=(0, 5))
        main_entry = ttk.Entry(fields_frame, textvariable=self.main_email_var,
                               font=ModernTheme.FONT_NORMAL)
        main_entry.pack(fill=tk.X, pady=(0, 15))

        # Sección de CCs
        ttk.Label(fields_frame, text="📋 Copias (CC) - Separar por comas:",
                 font=ModernTheme.FONT_NORMAL).pack(anchor=tk.W, pady=(0, 5))

        # Text widget para CC (con scroll)
        cc_frame = ttk.Frame(fields_frame)
        cc_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        cc_scrollbar = ttk.Scrollbar(cc_frame)
        cc_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.cc_text = tk.Text(cc_frame, height=6, font=ModernTheme.FONT_NORMAL,
                               yscrollcommand=cc_scrollbar.set, wrap=tk.WORD)
        self.cc_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        cc_scrollbar.config(command=self.cc_text.yview)

        # Nota informativa
        note_frame = tk.Frame(fields_frame, bg=ModernTheme.INFO,
                             highlightbackground=ModernTheme.SECONDARY,
                             highlightthickness=1)
        note_frame.pack(fill=tk.X, pady=(5, 0))

        note_label = tk.Label(note_frame,
                             text="💡 Ingrese los emails separados por comas. Ejemplo:\nemail1@example.com, email2@example.com",
                             fg=ModernTheme.TEXT_WHITE, bg=ModernTheme.INFO,
                             font=ModernTheme.FONT_SMALL, justify=tk.LEFT, pady=8, padx=10)
        note_label.pack()

        # Estado
        self.status_label = tk.Label(fields_frame, text="",
                                     font=ModernTheme.FONT_SMALL,
                                     fg=ModernTheme.TEXT_SECONDARY)
        self.status_label.pack(pady=(10, 0))

        # Frame de botones
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, side=tk.BOTTOM)

        # Botón guardar
        save_btn = ttk.Button(buttons_frame, text="💾 Guardar",
                             command=self.save_config,
                             style="Success.TButton")
        save_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5), ipady=8)

        # Botón cancelar
        cancel_btn = ttk.Button(buttons_frame, text="❌ Cancelar",
                               command=self.window.destroy,
                               style="TButton")
        cancel_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 0), ipady=8)


    def load_existing_config(self):
        """Carga configuración existente."""
        try:
            config = self.config_manager.load_config()
            if config:
                recipients_config = config.get("recipients_config")
                if recipients_config:
                    # Cargar destinatario principal
                    main_recipient = recipients_config.get("main_recipient", "")
                    self.main_email_var.set(main_recipient)

                    # Cargar CCs en el Text widget (convertir lista a texto separado por comas)
                    cc_recipients = recipients_config.get("cc_recipients", [])
                    if cc_recipients:
                        cc_text = ", ".join(cc_recipients)
                        self.cc_text.delete("1.0", tk.END)
                        self.cc_text.insert("1.0", cc_text)

                    if main_recipient:
                        self.update_status("🟡 Configuración cargada", "orange")
        except Exception as e:
            print(f"Error cargando configuración: {e}")

    def save_config(self):
        """Guarda la configuración de destinatarios."""
        recipients_data = self._get_recipients_data()
        is_valid, error_msg = self._validate_recipients_data(recipients_data)

        if not is_valid:
            self.update_status(f"🔴 {error_msg}", "red")
            messagebox.showerror("Error de validación", error_msg)
            return

        try:
            # Cargar configuración existente para no sobreescribir otros datos
            existing_config = self.config_manager.load_config() or {}

            # Actualizar solo los destinatarios
            existing_config["recipients_config"] = recipients_data

            # Guardar
            self.config_manager.save_config(existing_config)

            self.update_status("🟢 Configuración guardada", "green")
            messagebox.showinfo("Éxito", "Configuración de destinatarios guardada correctamente")

        except Exception as e:
            self.update_status(f"🔴 Error: {str(e)}", "red")
            messagebox.showerror("Error", f"Error al guardar configuración: {str(e)}")

    def _get_recipients_data(self):
        """Obtiene los datos de destinatarios actuales."""
        # Obtener texto del widget Text y parsear emails separados por comas
        cc_text = self.cc_text.get("1.0", tk.END).strip()

        # Parsear emails separados por comas
        cc_recipients = []
        if cc_text:
            # Dividir por comas y limpiar espacios
            cc_recipients = [email.strip() for email in cc_text.split(',')
                           if email.strip()]

        return {
            'main_recipient': self.main_email_var.get().strip(),
            'cc_recipients': cc_recipients
        }

    def _validate_recipients_data(self, recipients_data):
        """Valida los datos de destinatarios."""
        main_email = recipients_data.get('main_recipient', '').strip()
        cc_emails = recipients_data.get('cc_recipients', [])

        if not main_email:
            return False, "El destinatario principal es obligatorio"

        if not self._validate_email_format(main_email):
            return False, "El formato del destinatario principal es inválido"

        # Validar CCs
        for i, cc_email in enumerate(cc_emails):
            if cc_email.strip() and not self._validate_email_format(cc_email):
                return False, f"El formato del CC #{i + 1} es inválido"

        # Verificar duplicados
        all_emails = [main_email] + [cc for cc in cc_emails if cc.strip()]
        if len(all_emails) != len(set(email.lower() for email in all_emails)):
            return False, "Hay emails duplicados en la configuración"

        return True, "Configuración válida"

    def _validate_email_format(self, email):
        """Valida formato de email."""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, email.strip()))

    def update_status(self, message, color):
        """Actualiza el estado con colores modernos."""
        if self.status_label:
            color_map = {
                'green': ModernTheme.SUCCESS,
                'red': ModernTheme.DANGER,
                'orange': ModernTheme.WARNING,
                'blue': ModernTheme.INFO
            }
            modern_color = color_map.get(color, color)
            self.status_label.config(text=message, fg=modern_color)


class SearchConfigModal:
    """Modal para configuración de búsqueda de correos."""

    def __init__(self, parent):
        """Inicializa el modal de configuración de búsqueda."""
        self.parent = parent
        self.config_manager = ConfigManager()

        # Variables de búsqueda
        self.download_folder_var = tk.StringVar()

        # Crear ventana modal
        self.window = tk.Toplevel(parent)
        self.window.title("🔍 Configuración de Búsqueda")
        self.window.geometry("550x350")
        self.window.resizable(False, False)

        # Hacer modal (bloquear ventana principal)
        self.window.transient(parent)
        self.window.grab_set()

        # Centrar ventana
        self._center_window()

        # Crear interfaz
        self.create_interface()

        # Cargar configuración existente
        self.load_existing_config()

    def _center_window(self):
        """Centra la ventana modal en la pantalla."""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")

    def create_interface(self):
        """Crea la interfaz del modal."""
        # Frame principal con padding
        main_frame = ttk.Frame(self.window, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Título
        title_label = tk.Label(main_frame, text="🔍 Configuración de Búsqueda",
                              font=ModernTheme.FONT_HEADING,
                              fg=ModernTheme.PRIMARY)
        title_label.pack(pady=(0, 20))

        # Frame para campos
        fields_frame = ttk.Frame(main_frame)
        fields_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        # Carpeta de descarga
        ttk.Label(fields_frame, text="📁 Carpeta de Descarga:",
                 font=ModernTheme.FONT_NORMAL).pack(anchor=tk.W, pady=(0, 5))

        folder_frame = ttk.Frame(fields_frame)
        folder_frame.pack(fill=tk.X, pady=(0, 15))

        folder_entry = ttk.Entry(folder_frame, textvariable=self.download_folder_var,
                                font=ModernTheme.FONT_NORMAL)
        folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        browse_btn = ttk.Button(folder_frame, text="📁 Buscar",
                               command=self.browse_folder,
                               style="TButton")
        browse_btn.pack(side=tk.RIGHT)

        # Nota informativa
        note_frame = tk.Frame(fields_frame, bg=ModernTheme.INFO,
                             highlightbackground=ModernTheme.SECONDARY,
                             highlightthickness=1)
        note_frame.pack(fill=tk.X, pady=(10, 0))

        note_label = tk.Label(note_frame,
                             text="💡 Búsqueda fija: 'Cargador' con archivos Excel del día",
                             fg=ModernTheme.TEXT_WHITE, bg=ModernTheme.INFO,
                             font=ModernTheme.FONT_SMALL, justify=tk.CENTER, pady=8)
        note_label.pack()

        # Estado
        self.status_label = tk.Label(fields_frame, text="",
                                     font=ModernTheme.FONT_SMALL,
                                     fg=ModernTheme.TEXT_SECONDARY)
        self.status_label.pack(pady=(10, 0))

        # Frame de botones
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, side=tk.BOTTOM)

        # Botón restaurar
        restore_btn = ttk.Button(buttons_frame, text="🔄 Restaurar",
                                command=self.set_default_folder,
                                style="TButton")
        restore_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5), ipady=8)

        # Botón guardar
        save_btn = ttk.Button(buttons_frame, text="💾 Guardar",
                             command=self.save_config,
                             style="Success.TButton")
        save_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=5, ipady=8)

        # Botón cancelar
        cancel_btn = ttk.Button(buttons_frame, text="❌ Cancelar",
                               command=self.window.destroy,
                               style="TButton")
        cancel_btn.pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 0), ipady=8)

    def browse_folder(self):
        """Abre diálogo para seleccionar carpeta."""
        folder_path = filedialog.askdirectory(
            title="Seleccionar carpeta de descarga",
            initialdir=os.path.expanduser("~")
        )
        if folder_path:
            self.download_folder_var.set(folder_path)
            self.update_status("🟡 Carpeta seleccionada. Guarde la configuración.", "orange")

    def set_default_folder(self):
        """Establece la carpeta por defecto."""
        default_folder = os.path.join(os.path.expanduser("~"), "Downloads", "ContaFlow_Cargador")
        self.download_folder_var.set(default_folder)
        self.update_status("🟡 Carpeta por defecto establecida. Guarde la configuración.", "orange")

    def load_existing_config(self):
        """Carga configuración existente."""
        try:
            config = self.config_manager.load_config()
            if config:
                search_criteria = config.get("search_criteria", {})
                folder = search_criteria.get("download_folder", "")

                if folder:
                    self.download_folder_var.set(folder)
                    self.update_status("🟡 Configuración cargada", "orange")
                else:
                    self.set_default_folder()
        except Exception as e:
            print(f"Error cargando configuración: {e}")
            self.set_default_folder()

    def save_config(self):
        """Guarda la configuración de búsqueda."""
        download_folder = self.download_folder_var.get().strip()

        if not download_folder:
            self.update_status("🔴 Debe especificar una carpeta de descarga", "red")
            messagebox.showerror("Error", "Debe especificar una carpeta de descarga")
            return

        try:
            # Crear carpeta si no existe
            os.makedirs(download_folder, exist_ok=True)

            # Cargar configuración existente
            existing_config = self.config_manager.load_config() or {}

            # Actualizar solo la configuración de búsqueda
            config_data = {
                "subject": "Cargador",
                "search_type": "Contiene",
                "today_only": True,
                "attachments_only": True,
                "excel_files": True,
                "download_folder": download_folder
            }

            existing_config["search_criteria"] = config_data

            # Guardar
            self.config_manager.save_config(existing_config)

            self.update_status("🟢 Configuración guardada", "green")
            messagebox.showinfo("Éxito", "Configuración de búsqueda guardada correctamente")

        except Exception as e:
            self.update_status(f"🔴 Error: {str(e)}", "red")
            messagebox.showerror("Error", f"Error al guardar configuración: {str(e)}")

    def update_status(self, message, color):
        """Actualiza el estado con colores modernos."""
        if self.status_label:
            color_map = {
                'green': ModernTheme.SUCCESS,
                'red': ModernTheme.DANGER,
                'orange': ModernTheme.WARNING,
                'blue': ModernTheme.INFO
            }
            modern_color = color_map.get(color, color)
            self.status_label.config(text=message, fg=modern_color)


class XmlConfigModal:
    """Modal para configuración de procesamiento XML."""

    def __init__(self, parent):
        """Inicializa el modal de configuración XML."""
        self.parent = parent
        self.config_manager = ConfigManager()

        # Definir empresas con sus campos
        self.company_folders = {
            "nargallo": {
                "name": "Nargallo del Este S.A.",
                "folder_var": tk.StringVar(),
                "activity_var": tk.StringVar()
            },
            "ventas_fruno": {
                "name": "Ventas Fruno, S.A.",
                "folder_var": tk.StringVar(),
                "activity_var": tk.StringVar()
            },
            "creme_caramel": {
                "name": "Creme Caramel",
                "folder_var": tk.StringVar(),
                "activity_var": tk.StringVar()
            },
            "su_laka": {
                "name": "Su Laka",
                "folder_var": tk.StringVar(),
                "activity_var": tk.StringVar()
            }
        }

        # Variables adicionales
        self.output_folder_var = tk.StringVar()
        self.delete_originals_var = tk.BooleanVar(value=True)
        self.auto_send_var = tk.BooleanVar(value=True)

        # Crear ventana modal
        self.window = tk.Toplevel(parent)
        self.window.title("📄 Configuración XML")
        self.window.geometry("700x600")
        self.window.resizable(False, False)

        # Hacer modal
        self.window.transient(parent)
        self.window.grab_set()

        # Centrar ventana
        self._center_window()

        # Crear interfaz
        self.create_interface()

        # Cargar configuración existente
        self.load_existing_config()

    def _center_window(self):
        """Centra la ventana modal en la pantalla."""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")

    def create_interface(self):
        """Crea la interfaz del modal."""
        # Frame principal con padding
        main_frame = ttk.Frame(self.window, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Título
        title_label = tk.Label(main_frame, text="📄 Configuración XML",
                              font=ModernTheme.FONT_HEADING,
                              fg=ModernTheme.PRIMARY)
        title_label.pack(pady=(0, 10))

        # Mensaje informativo
        month_info = self.config_manager.get_current_month_folder_info()
        info_text = f"💡 Configure carpetas BASE. Se agregará automáticamente \\{month_info['folder_suffix']}"

        info_frame = tk.Frame(main_frame, bg=ModernTheme.INFO,
                             highlightbackground=ModernTheme.SECONDARY,
                             highlightthickness=1)
        info_frame.pack(fill=tk.X, pady=(0, 10))

        tk.Label(info_frame, text=info_text,
                fg=ModernTheme.TEXT_WHITE, bg=ModernTheme.INFO,
                font=ModernTheme.FONT_SMALL, pady=5, padx=10).pack()

        # Frame scrollable para empresas
        companies_frame = ttk.LabelFrame(main_frame, text="🗂️ Carpetas por Empresa", padding=10)
        companies_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Canvas con scroll
        canvas = tk.Canvas(companies_frame, height=250, highlightthickness=0)
        scrollbar = ttk.Scrollbar(companies_frame, orient="vertical", command=canvas.yview)
        scroll_frame = ttk.Frame(canvas)

        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        canvas.configure(yscrollcommand=scrollbar.set)
        canvas_window = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

        # Crear campos para cada empresa
        for company_key, company_info in self.company_folders.items():
            self.create_company_fields(scroll_frame, company_key, company_info)

        # Configurar scroll
        def configure_scroll(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))
            canvas_width = canvas.winfo_width()
            if canvas_width > 1:
                canvas.itemconfig(canvas_window, width=canvas_width)

        scroll_frame.bind("<Configure>", configure_scroll)
        canvas.bind("<Configure>", configure_scroll)

        # Carpeta de salida
        output_frame = ttk.LabelFrame(main_frame, text="💾 Carpeta de Salida", padding=10)
        output_frame.pack(fill=tk.X, pady=(0, 10))

        folder_frame = ttk.Frame(output_frame)
        folder_frame.pack(fill=tk.X)

        ttk.Entry(folder_frame, textvariable=self.output_folder_var,
                 font=ModernTheme.FONT_NORMAL).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(folder_frame, text="📁", width=3,
                  command=self.browse_output_folder).pack(side=tk.RIGHT)

        # Opciones
        options_frame = ttk.LabelFrame(main_frame, text="⚙️ Opciones", padding=10)
        options_frame.pack(fill=tk.X, pady=(0, 10))

        ttk.Checkbutton(options_frame, text="Eliminar archivos Excel originales",
                       variable=self.delete_originals_var).pack(anchor=tk.W)
        ttk.Checkbutton(options_frame, text="Enviar automáticamente",
                       variable=self.auto_send_var).pack(anchor=tk.W)

        # Estado
        self.status_label = tk.Label(main_frame, text="",
                                     font=ModernTheme.FONT_SMALL,
                                     fg=ModernTheme.TEXT_SECONDARY)
        self.status_label.pack(pady=(5, 10))

        # Botones
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, side=tk.BOTTOM)

        ttk.Button(buttons_frame, text="💾 Guardar",
                  command=self.save_config,
                  style="Success.TButton").pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5), ipady=8)

        ttk.Button(buttons_frame, text="❌ Cancelar",
                  command=self.window.destroy,
                  style="TButton").pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 0), ipady=8)

    def create_company_fields(self, parent, company_key, company_info):
        """Crea campos para una empresa."""
        company_frame = ttk.LabelFrame(parent, text=f"📁 {company_info['name']}", padding=5)
        company_frame.pack(fill=tk.X, pady=3)

        # Carpeta BASE
        ttk.Label(company_frame, text="Carpeta BASE:", font=("Arial", 9)).pack(anchor=tk.W)

        folder_frame = ttk.Frame(company_frame)
        folder_frame.pack(fill=tk.X, pady=2)

        ttk.Entry(folder_frame, textvariable=company_info['folder_var'],
                 font=("Arial", 9)).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(folder_frame, text="📁", width=3,
                  command=lambda: self.browse_company_folder(company_info)).pack(side=tk.RIGHT)

        # Actividad comercial
        ttk.Label(company_frame, text="Actividad comercial (opcional):",
                 font=("Arial", 9)).pack(anchor=tk.W, pady=(5, 2))
        ttk.Entry(company_frame, textvariable=company_info['activity_var'],
                 font=("Arial", 9)).pack(fill=tk.X)

    def browse_company_folder(self, company_info):
        """Abre diálogo para seleccionar carpeta de empresa."""
        folder_path = filedialog.askdirectory(
            title=f"Seleccionar carpeta BASE para {company_info['name']}",
            initialdir=os.path.expanduser("~")
        )
        if folder_path:
            company_info['folder_var'].set(folder_path)

    def browse_output_folder(self):
        """Abre diálogo para seleccionar carpeta de salida."""
        folder_path = filedialog.askdirectory(
            title="Seleccionar carpeta para archivos procesados",
            initialdir=os.path.expanduser("~")
        )
        if folder_path:
            self.output_folder_var.set(folder_path)

    def load_existing_config(self):
        """Carga configuración existente."""
        try:
            config = self.config_manager.load_config()

            if config and "xml_config" in config:
                xml_config = config["xml_config"]

                # Cargar carpetas empresariales
                company_folders = xml_config.get("company_folders", {})
                for company_key, company_info in self.company_folders.items():
                    if company_key in company_folders:
                        company_info['folder_var'].set(company_folders[company_key])

                # Cargar actividades comerciales
                commercial_activities = xml_config.get("commercial_activities", {})
                for company_key, company_info in self.company_folders.items():
                    if company_key in commercial_activities:
                        company_info['activity_var'].set(commercial_activities[company_key])

                # Otras configuraciones
                self.output_folder_var.set(xml_config.get("output_folder", ""))
                self.delete_originals_var.set(xml_config.get("delete_originals", True))
                self.auto_send_var.set(xml_config.get("auto_send", True))

                configured_count = len([k for k, v in company_folders.items() if v])
                if configured_count > 0:
                    self.update_status(f"🟡 Configuración cargada: {configured_count} carpetas", "orange")
            else:
                # Valor por defecto para carpeta de salida
                default_output = os.path.join(os.path.expanduser("~"), "Downloads", "ContaFlow", "Procesados")
                self.output_folder_var.set(default_output)

        except Exception as e:
            print(f"Error cargando configuración XML: {e}")

    def save_config(self):
        """Guarda la configuración XML."""
        # Obtener carpetas configuradas
        company_folders = {}
        commercial_activities = {}

        for company_key, company_info in self.company_folders.items():
            folder_path = company_info['folder_var'].get().strip()
            if folder_path:
                company_folders[company_key] = folder_path

            activity = company_info['activity_var'].get().strip()
            commercial_activities[company_key] = activity

        output_folder = self.output_folder_var.get().strip()

        if not company_folders:
            self.update_status("🔴 Configure al menos una carpeta", "red")
            messagebox.showerror("Error", "Configure al menos una carpeta de empresa")
            return

        if not output_folder:
            self.update_status("🔴 Especifique carpeta de salida", "red")
            messagebox.showerror("Error", "Especifique la carpeta de salida")
            return

        # Validar carpetas BASE
        for company_key, base_folder_path in company_folders.items():
            if not os.path.exists(base_folder_path):
                company_name = self.company_folders[company_key]["name"]
                self.update_status(f"🔴 Carpeta no existe: {company_name}", "red")
                messagebox.showerror("Error", f"La carpeta no existe: {company_name}")
                return

        try:
            # Crear carpeta de salida si no existe
            os.makedirs(output_folder, exist_ok=True)

            # Cargar configuración existente
            existing_config = self.config_manager.load_config() or {}

            # Crear configuración XML
            xml_config = {
                "company_folders": company_folders,
                "commercial_activities": commercial_activities,
                "output_folder": output_folder,
                "delete_originals": self.delete_originals_var.get(),
                "auto_send": self.auto_send_var.get(),
                "detailed_logs": True,
                "manual_review_limit": 3
            }

            # Preservar exclusiones de combustible si existen
            if "xml_config" in existing_config and "combustible_exclusions" in existing_config["xml_config"]:
                xml_config["combustible_exclusions"] = existing_config["xml_config"]["combustible_exclusions"]

            existing_config["xml_config"] = xml_config

            # Guardar
            self.config_manager.save_config(existing_config)

            configured_count = len(company_folders)
            self.update_status(f"🟢 Guardado: {configured_count} carpetas configuradas", "green")
            messagebox.showinfo("Éxito", "Configuración XML guardada correctamente")

        except Exception as e:
            self.update_status(f"🔴 Error: {str(e)}", "red")
            messagebox.showerror("Error", f"Error al guardar configuración: {str(e)}")

    def update_status(self, message, color):
        """Actualiza el estado con colores modernos."""
        if self.status_label:
            color_map = {
                'green': ModernTheme.SUCCESS,
                'red': ModernTheme.DANGER,
                'orange': ModernTheme.WARNING,
                'blue': ModernTheme.INFO
            }
            modern_color = color_map.get(color, color)
            self.status_label.config(text=message, fg=modern_color)


class CombustibleExclusionsModal:
    """Modal para configuración de exclusiones de combustible."""

    def __init__(self, parent):
        """Inicializa el modal de exclusiones de combustible."""
        self.parent = parent
        self.config_manager = ConfigManager()

        # Variables
        self.emitter_var = tk.StringVar()
        self.exclusions: List[str] = []
        self.xml_config_available = False

        # Crear ventana modal
        self.window = tk.Toplevel(parent)
        self.window.title("⛽ Exclusiones de Combustible")
        self.window.geometry("550x500")
        self.window.resizable(False, False)

        # Hacer modal
        self.window.transient(parent)
        self.window.grab_set()

        # Centrar ventana
        self._center_window()

        # Crear interfaz
        self.create_interface()

        # Cargar configuración existente
        self.load_existing_config()

    def _center_window(self):
        """Centra la ventana modal en la pantalla."""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f"{width}x{height}+{x}+{y}")

    def create_interface(self):
        """Crea la interfaz del modal."""
        # Frame principal con padding
        main_frame = ttk.Frame(self.window, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Título
        title_label = tk.Label(main_frame, text="⛽ Exclusiones de Combustible",
                              font=ModernTheme.FONT_HEADING,
                              fg=ModernTheme.PRIMARY)
        title_label.pack(pady=(0, 10))

        # Descripción
        desc_frame = tk.Frame(main_frame, bg=ModernTheme.INFO,
                             highlightbackground=ModernTheme.SECONDARY,
                             highlightthickness=1)
        desc_frame.pack(fill=tk.X, pady=(0, 15))

        description = (
            "💬 Agrega los valores exactos de <NombreEmisor> que deben excluirse del "
            "tratamiento como Combustible/Placa. Cuando se detecte un emisor en "
            "esta lista, el bot utilizará el comportamiento normal."
        )
        tk.Label(desc_frame, text=description, wraplength=480, justify=tk.LEFT,
                bg=ModernTheme.INFO, fg=ModernTheme.TEXT_WHITE,
                font=ModernTheme.FONT_SMALL, pady=8, padx=10).pack()

        # Frame para agregar
        add_frame = ttk.LabelFrame(main_frame, text="➕ Agregar Exclusión", padding=10)
        add_frame.pack(fill=tk.X, pady=(0, 10))

        input_frame = ttk.Frame(add_frame)
        input_frame.pack(fill=tk.X)

        ttk.Entry(input_frame, textvariable=self.emitter_var,
                 font=ModernTheme.FONT_NORMAL).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(input_frame, text="Agregar",
                  command=self.add_exclusion,
                  style="Primary.TButton").pack(side=tk.RIGHT, ipady=4)

        # Lista de exclusiones
        list_frame = ttk.LabelFrame(main_frame, text="📋 Emisores Excluidos", padding=10)
        list_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Listbox con scrollbar
        listbox_frame = ttk.Frame(list_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True)

        self.listbox = tk.Listbox(listbox_frame, height=8,
                                  font=ModernTheme.FONT_NORMAL,
                                  bg=ModernTheme.BG_SURFACE,
                                  fg=ModernTheme.TEXT_PRIMARY,
                                  selectbackground=ModernTheme.SECONDARY,
                                  selectforeground=ModernTheme.TEXT_WHITE)
        self.listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(listbox_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.listbox.configure(yscrollcommand=scrollbar.set)

        # Botones de lista
        list_buttons = ttk.Frame(list_frame)
        list_buttons.pack(fill=tk.X, pady=(5, 0))

        ttk.Button(list_buttons, text="Eliminar seleccionado",
                  command=self.remove_selected,
                  style="TButton").pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))

        ttk.Button(list_buttons, text="Limpiar lista",
                  command=self.clear_exclusions,
                  style="TButton").pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 0))

        # Estado
        self.status_label = tk.Label(main_frame, text="",
                                     font=ModernTheme.FONT_SMALL,
                                     fg=ModernTheme.TEXT_SECONDARY)
        self.status_label.pack(pady=(5, 10))

        # Botones principales
        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.pack(fill=tk.X, side=tk.BOTTOM)

        ttk.Button(buttons_frame, text="💾 Guardar",
                  command=self.save_config,
                  style="Success.TButton").pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5), ipady=8)

        ttk.Button(buttons_frame, text="❌ Cancelar",
                  command=self.window.destroy,
                  style="TButton").pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(5, 0), ipady=8)

    def load_existing_config(self):
        """Carga configuración existente."""
        try:
            config = self.config_manager.load_config() or {}
            xml_config = config.get('xml_config')

            if isinstance(xml_config, dict) and (
                xml_config.get('company_folders') or xml_config.get('xml_folder')
            ):
                self.xml_config_available = True
                exclusions_config = xml_config.get('combustible_exclusions', {})

                if isinstance(exclusions_config, dict):
                    emitter_names = exclusions_config.get('emitter_names', [])
                elif isinstance(exclusions_config, list):
                    emitter_names = exclusions_config
                else:
                    emitter_names = []

                self.exclusions = [name for name in emitter_names if isinstance(name, str) and name.strip()]
                self.exclusions.sort(key=lambda x: x.lower())
                self._refresh_listbox()

                if self.exclusions:
                    self.update_status("🟢 Exclusiones cargadas correctamente", "green")
                else:
                    self.update_status("🟡 No hay exclusiones configuradas", "orange")
            else:
                self.xml_config_available = False
                self.exclusions = []
                self._refresh_listbox()
                self.update_status("🔴 Debe configurar XML antes de agregar exclusiones", "red")
        except Exception as e:
            self.update_status(f"🔴 Error al cargar: {str(e)}", "red")

    def add_exclusion(self):
        """Agrega una exclusión."""
        name = self.emitter_var.get().strip()
        if not name:
            self.update_status("🔴 Debe ingresar un valor para el emisor", "red")
            return

        normalized = self._normalize_name(name)
        if normalized in {self._normalize_name(item) for item in self.exclusions}:
            self.update_status("⚠️ El emisor ya se encuentra en la lista", "orange")
            return

        self.exclusions.append(name)
        self.exclusions.sort(key=lambda x: x.lower())
        self._refresh_listbox()
        self.emitter_var.set("")
        self.update_status("🟢 Emisor agregado a las exclusiones", "green")

    def remove_selected(self):
        """Elimina la exclusión seleccionada."""
        if not self.listbox.curselection():
            self.update_status("⚠️ Debe seleccionar un emisor para eliminar", "orange")
            return

        index = self.listbox.curselection()[0]
        removed = self.exclusions.pop(index)
        self._refresh_listbox()
        self.update_status(f"🟢 Emisor eliminado: {removed}", "green")

    def clear_exclusions(self):
        """Limpia todas las exclusiones."""
        if not self.exclusions:
            self.update_status("⚠️ La lista de exclusiones ya está vacía", "orange")
            return

        self.exclusions.clear()
        self._refresh_listbox()
        self.update_status("🟡 Lista limpiada. Recuerde guardar los cambios", "orange")

    def save_config(self):
        """Guarda la configuración de exclusiones."""
        if not self.xml_config_available:
            self.update_status("🔴 Configure XML antes de guardar exclusiones", "red")
            messagebox.showerror("Error", "Debe configurar XML antes de guardar exclusiones")
            return

        try:
            config = self.config_manager.load_config() or {}
            xml_config = config.get('xml_config')

            if not isinstance(xml_config, dict):
                self.update_status("🔴 Configuración XML inválida", "red")
                return

            # Actualizar exclusiones
            xml_config['combustible_exclusions'] = {
                'emitter_names': self.exclusions
            }

            config['xml_config'] = xml_config

            # Guardar
            self.config_manager.save_config(config)
            self.update_status("🟢 Exclusiones guardadas correctamente", "green")
            messagebox.showinfo("Éxito", "Exclusiones guardadas correctamente")

        except Exception as e:
            self.update_status(f"🔴 Error al guardar: {str(e)}", "red")
            messagebox.showerror("Error", f"Error al guardar: {str(e)}")

    def _refresh_listbox(self):
        """Actualiza el listbox."""
        if not self.listbox:
            return
        self.listbox.delete(0, tk.END)
        for item in self.exclusions:
            self.listbox.insert(tk.END, item)

    def update_status(self, message, color):
        """Actualiza el estado con colores modernos."""
        if self.status_label:
            color_map = {
                'green': ModernTheme.SUCCESS,
                'red': ModernTheme.DANGER,
                'orange': ModernTheme.WARNING,
                'blue': ModernTheme.INFO
            }
            modern_color = color_map.get(color, color)
            self.status_label.config(text=message, fg=modern_color)

    @staticmethod
    def _normalize_name(name: str) -> str:
        """Normaliza un nombre para comparación."""
        if not name:
            return ""
        normalized = unicodedata.normalize('NFKD', name)
        normalized = ''.join(c for c in normalized if not unicodedata.combining(c))
        return normalized.strip().lower()
