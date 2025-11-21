# email_config_modals.py
"""
Modales para configuración de email y destinatarios.
Usados desde la pestaña de Automatización.
"""

import tkinter as tk
from tkinter import ttk, messagebox
import threading
import re
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

            # Cerrar modal después de 500ms
            self.window.after(500, self.window.destroy)

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
        self.cc_entries = []
        self.max_ccs = 10
        self.next_cc_id = 0

        # Crear ventana modal
        self.window = tk.Toplevel(parent)
        self.window.title("📧 Configurar Destinatarios")
        self.window.geometry("550x550")
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
        cc_header_frame = ttk.Frame(fields_frame)
        cc_header_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(cc_header_frame, text="📋 Copias (CC):",
                 font=ModernTheme.FONT_NORMAL).pack(side=tk.LEFT)
        self.cc_counter_label = ttk.Label(cc_header_frame, text=f"0/{self.max_ccs}",
                                          foreground=ModernTheme.TEXT_SECONDARY,
                                          font=ModernTheme.FONT_SMALL)
        self.cc_counter_label.pack(side=tk.RIGHT)

        # Frame con scroll para CCs
        self.create_cc_scroll_area(fields_frame)

        # Botón agregar CC
        self.add_cc_btn = ttk.Button(fields_frame, text="➕ Agregar CC",
                                     command=self.add_cc_field,
                                     style="TButton")
        self.add_cc_btn.pack(fill=tk.X, pady=(10, 0), ipady=4)

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

    def create_cc_scroll_area(self, parent):
        """Crea el área con scroll para los campos CC."""
        # Frame con canvas para scroll
        scroll_frame = ttk.Frame(parent)
        scroll_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))

        # Canvas y scrollbar
        self.canvas = tk.Canvas(scroll_frame, bg="white", height=200)
        scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=self.canvas.yview)

        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.canvas.configure(yscrollcommand=scrollbar.set)

        # Frame interno para los CCs
        self.cc_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.cc_frame, anchor="nw")

        # Configurar scroll
        def configure_scroll_region(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        self.cc_frame.bind("<Configure>", configure_scroll_region)

        # Mousewheel scroll
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        def _bind_mousewheel(event):
            self.canvas.bind("<MouseWheel>", _on_mousewheel)

        def _unbind_mousewheel(event):
            self.canvas.unbind("<MouseWheel>")

        self.canvas.bind("<Enter>", _bind_mousewheel)
        self.canvas.bind("<Leave>", _unbind_mousewheel)

    def add_cc_field(self):
        """Agrega un nuevo campo CC."""
        if len(self.cc_entries) >= self.max_ccs:
            self.update_status("⚠️ Máximo 10 CCs permitidos", "orange")
            return

        # Generar ID único
        cc_id = self.next_cc_id
        self.next_cc_id += 1

        # Frame para este CC
        cc_container = ttk.Frame(self.cc_frame)
        cc_container.pack(fill=tk.X, padx=5, pady=2)
        cc_container.grid_columnconfigure(0, weight=1)

        # Entry para email
        cc_entry = ttk.Entry(cc_container, font=("Arial", 9))
        cc_entry.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        # Botón eliminar
        remove_btn = ttk.Button(cc_container, text="❌", width=3,
                                command=lambda: self.remove_cc_field(cc_id))
        remove_btn.grid(row=0, column=1)

        # Guardar referencia
        cc_data = {
            'id': cc_id,
            'container': cc_container,
            'entry': cc_entry
        }
        self.cc_entries.append(cc_data)

        self.update_cc_counter()

    def remove_cc_field(self, cc_id):
        """Elimina un campo CC usando ID único."""
        cc_to_remove = None
        for cc in self.cc_entries:
            if cc['id'] == cc_id:
                cc_to_remove = cc
                break

        if not cc_to_remove:
            return

        # Destruir el container
        cc_to_remove['container'].destroy()

        # Remover de la lista
        self.cc_entries = [cc for cc in self.cc_entries if cc['id'] != cc_id]

        self.update_cc_counter()

    def update_cc_counter(self):
        """Actualiza el contador de CCs."""
        if self.cc_counter_label:
            self.cc_counter_label.config(text=f"{len(self.cc_entries)}/{self.max_ccs}")

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

                    # Cargar CCs
                    cc_recipients = recipients_config.get("cc_recipients", [])
                    for cc_email in cc_recipients:
                        if cc_email.strip():
                            self.add_cc_field()
                            # Establecer valor en el último CC agregado
                            if self.cc_entries:
                                self.cc_entries[-1]['entry'].insert(0, cc_email.strip())

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

            # Cerrar modal después de 500ms
            self.window.after(500, self.window.destroy)

        except Exception as e:
            self.update_status(f"🔴 Error: {str(e)}", "red")
            messagebox.showerror("Error", f"Error al guardar configuración: {str(e)}")

    def _get_recipients_data(self):
        """Obtiene los datos de destinatarios actuales."""
        return {
            'main_recipient': self.main_email_var.get().strip(),
            'cc_recipients': [cc['entry'].get().strip() for cc in self.cc_entries
                              if cc['entry'].get().strip()]
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
