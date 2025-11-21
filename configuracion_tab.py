# configuracion_tab.py
"""
Pestaña simplificada de configuración usando tkinter nativo.
Contiene sub-pestañas para email/destinatarios, búsqueda y XML únicamente.
"""
# Archivos relacionados: busqueda_tab.py, xml_tab.py, config_manager.py, email_manager.py

import tkinter as tk
from tkinter import ttk
import threading
import re
from email_manager import EmailManager
from config_manager import ConfigManager
from combustible_exclusions_tab import CombustibleExclusionsTab


class ConfiguracionTab:
    """Pestaña de configuración simplificada con 3 sub-pestañas optimizadas."""

    def __init__(self, parent):
        """Inicializa la pestaña de configuración."""
        self.parent = parent
        self.is_visible = False
        self.current_subtab = None

        # Managers
        self.email_manager = EmailManager()
        self.config_manager = ConfigManager()

        # Sub-pestañas
        self.subtabs = {}

        # Crear interfaz
        self.create_interface()
        self.initialize_subtabs()

        # Mostrar sub-pestaña por defecto
        self.show_subtab("email_destinatarios")

    def create_interface(self):
        """Crea la interfaz principal."""
        # Frame principal
        self.main_frame = ttk.Frame(self.parent)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(1, weight=1)

        # Título
        self.create_title()

        # Notebook para sub-pestañas (ahora con 4 pestañas)
        self.create_subtab_notebook()

    def create_title(self):
        """Crea el título de la pestaña con mejor diseño."""
        title_frame = ttk.LabelFrame(self.main_frame, text="Configuración del Sistema", padding=15)
        title_frame.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="ew")

        title_label = ttk.Label(
            title_frame,
            text="⚙️ Configuración del Sistema",
            font=("Arial", 14, "bold"),
            foreground="#2c3e50"
        )
        title_label.pack()

        subtitle_label = ttk.Label(
            title_frame,
            text="Gestiona credenciales, destinatarios, búsqueda y procesamiento XML",
            font=("Arial", 9),
            foreground="#7f8c8d"
        )
        subtitle_label.pack()

    def create_subtab_notebook(self):
        """Crea el notebook para sub-pestañas."""
        self.subtab_notebook = ttk.Notebook(self.main_frame)
        self.subtab_notebook.grid(row=1, column=0, padx=10, pady=5, sticky="nsew")

        # Crear frames para cada sub-pestaña (ahora 4 pestañas)
        self.email_destinatarios_frame = ttk.Frame(self.subtab_notebook)
        self.busqueda_frame = ttk.Frame(self.subtab_notebook)
        self.xml_frame = ttk.Frame(self.subtab_notebook)
        self.combustible_exclusions_frame = ttk.Frame(self.subtab_notebook)

        # Agregar pestañas al notebook
        self.subtab_notebook.add(self.email_destinatarios_frame, text="📧 Email y Destinatarios")
        self.subtab_notebook.add(self.busqueda_frame, text="🔍 Búsqueda")
        self.subtab_notebook.add(self.xml_frame, text="🗂️ XML")
        self.subtab_notebook.add(self.combustible_exclusions_frame, text="⛽️ Exclusiones Combustible")

        # Vincular evento de cambio
        self.subtab_notebook.bind("<<NotebookTabChanged>>", self._on_subtab_changed)

    def initialize_subtabs(self):
        """Inicializa las sub-pestañas."""
        try:
            # Crear sub-pestaña combinada de email y destinatarios
            self.subtabs["email_destinatarios"] = EmailDestinatariosSubTab(self.email_destinatarios_frame, self)
            print("✅ Sub-pestaña email y destinatarios inicializada")

            # Crear sub-pestaña de búsqueda
            from busqueda_tab import BusquedaTab
            self.subtabs["busqueda"] = BusquedaTab(self.busqueda_frame)
            print("✅ Sub-pestaña búsqueda inicializada")

            # Crear sub-pestaña de XML
            from xml_tab import XmlTab
            self.subtabs["xml"] = XmlTab(self.xml_frame)
            print("✅ Sub-pestaña XML inicializada")

            # Crear sub-pestaña de exclusiones de combustible
            self.subtabs["combustible_exclusiones"] = CombustibleExclusionsTab(
                self.combustible_exclusions_frame
            )
            print("✅ Sub-pestaña exclusiones de combustible inicializada")

        except Exception as e:
            print(f"⚠️ Error inicializando sub-pestañas: {e}")

    def _on_subtab_changed(self, event):
        """Maneja cambio de sub-pestaña."""
        try:
            selected_index = event.widget.index('current')
            subtab_names = ["email_destinatarios", "busqueda", "xml", "combustible_exclusiones"]

            if 0 <= selected_index < len(subtab_names):
                self.show_subtab(subtab_names[selected_index])

        except Exception as e:
            print(f"⚠️ Error en cambio de sub-pestaña: {e}")

    def show_subtab(self, subtab_name):
        """Muestra la sub-pestaña especificada."""
        try:
            if subtab_name not in self.subtabs:
                print(f"⚠️ Sub-pestaña desconocida: {subtab_name}")
                return

            if self.current_subtab == subtab_name:
                return

            # Ocultar sub-pestaña actual
            if self.current_subtab and self.subtabs.get(self.current_subtab):
                if hasattr(self.subtabs[self.current_subtab], 'hide'):
                    self.subtabs[self.current_subtab].hide()

            # Mostrar nueva sub-pestaña
            if hasattr(self.subtabs[subtab_name], 'show'):
                self.subtabs[subtab_name].show()
                self.current_subtab = subtab_name

                # Cargar configuración si es email_destinatarios
                if subtab_name == "email_destinatarios":
                    self.load_existing_config()

        except Exception as e:
            print(f"⚠️ Error mostrando sub-pestaña {subtab_name}: {e}")

    def load_existing_config(self):
        """Carga configuración existente para la pestaña combinada."""
        if self.current_subtab == "email_destinatarios" and self.subtabs["email_destinatarios"]:
            try:
                self.subtabs["email_destinatarios"].load_existing_config()
            except Exception as e:
                print(f"⚠️ Error cargando configuración: {e}")

    def show(self):
        """Muestra la pestaña de configuración."""
        if not self.is_visible:
            try:
                self.main_frame.pack(fill=tk.BOTH, expand=True)
                self.is_visible = True

                # Asegurar que sub-pestaña actual esté visible
                if self.current_subtab and self.subtabs.get(self.current_subtab):
                    self.subtabs[self.current_subtab].show()

            except Exception as e:
                print(f"Error mostrando configuración: {e}")

    def hide(self):
        """Oculta la pestaña de configuración."""
        if self.is_visible:
            try:
                # Ocultar sub-pestaña actual
                if self.current_subtab and self.subtabs.get(self.current_subtab):
                    if hasattr(self.subtabs[self.current_subtab], 'hide'):
                        self.subtabs[self.current_subtab].hide()

                self.main_frame.pack_forget()
                self.is_visible = False

            except Exception as e:
                print(f"Error ocultando configuración: {e}")

    def get_current_subtab(self):
        """Obtiene la sub-pestaña actual."""
        return self.current_subtab


class EmailDestinatariosSubTab:
    """Sub-pestaña combinada para credenciales de email y configuración de destinatarios con diseño de 3 columnas."""

    def __init__(self, parent, config_tab):
        """Inicializa la sub-pestaña combinada."""
        self.parent = parent
        self.config_tab = config_tab
        self.is_visible = False

        # Variables de credenciales
        self.provider_var = tk.StringVar(value="Gmail")
        self.email_var = tk.StringVar()
        self.password_var = tk.StringVar()

        # Variables de destinatarios
        self.main_email_var = tk.StringVar()
        self.cc_entries = []
        self.max_ccs = 10

        # Referencias a widgets
        self.credentials_status_label = None
        self.recipients_status_label = None
        self.cc_frame = None
        self.cc_counter_label = None
        self.add_cc_btn = None

        # Crear interfaz
        self.create_interface()

    def create_interface(self):
        """Crea la interfaz con diseño de 3 columnas."""
        # Frame principal con 3 columnas
        main_frame = ttk.Frame(self.parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Configurar grid de 3 columnas
        main_frame.grid_columnconfigure(0, weight=1)  # Email
        main_frame.grid_columnconfigure(1, weight=1)  # Destinatarios
        main_frame.grid_columnconfigure(2, weight=1)  # Control
        main_frame.grid_rowconfigure(0, weight=1)

        # Columna 1: Configuración de Email
        self.create_email_column(main_frame)

        # Columna 2: Configuración de Destinatarios
        self.create_recipients_column(main_frame)

        # Columna 3: Control y Estado
        self.create_control_column(main_frame)

    def create_email_column(self, parent):
        """Crea la columna de configuración de email."""
        email_frame = ttk.LabelFrame(parent, text="📧 Configuración de Email", padding=15)
        email_frame.grid(row=0, column=0, padx=(0, 5), pady=0, sticky="nsew")

        # Proveedor
        ttk.Label(email_frame, text="Proveedor de Correo:").pack(anchor=tk.W, pady=(0, 5))
        provider_combo = ttk.Combobox(email_frame, textvariable=self.provider_var,
                                      values=["Gmail", "Outlook", "Yahoo", "Otro"],
                                      state="readonly", width=30)
        provider_combo.pack(fill=tk.X, pady=(0, 15))

        # Email
        ttk.Label(email_frame, text="Correo Electrónico:").pack(anchor=tk.W, pady=(0, 5))
        email_entry = ttk.Entry(email_frame, textvariable=self.email_var, width=35)
        email_entry.pack(fill=tk.X, pady=(0, 15))

        # Contraseña
        ttk.Label(email_frame, text="Contraseña:").pack(anchor=tk.W, pady=(0, 5))
        password_entry = ttk.Entry(email_frame, textvariable=self.password_var,
                                   show="*", width=35)
        password_entry.pack(fill=tk.X, pady=(0, 15))

        # Nota informativa
        note_frame = ttk.Frame(email_frame)
        note_frame.pack(fill=tk.X, pady=(10, 0))

        note_label = ttk.Label(note_frame, text="💡 Para Gmail usa una\ncontraseña de aplicación",
                               foreground="blue", font=("Arial", 9), justify=tk.CENTER)
        note_label.pack()

        # Espaciador para empujar contenido hacia arriba
        spacer = ttk.Frame(email_frame)
        spacer.pack(fill=tk.BOTH, expand=True)

    def create_recipients_column(self, parent):
        """Crea la columna de configuración de destinatarios."""
        recipients_frame = ttk.LabelFrame(parent, text="👥 Configuración de Destinatarios", padding=15)
        recipients_frame.grid(row=0, column=1, padx=5, pady=0, sticky="nsew")
        recipients_frame.grid_rowconfigure(2, weight=1)

        # Destinatario principal
        ttk.Label(recipients_frame, text="📧 Destinatario Principal:").pack(anchor=tk.W, pady=(0, 5))
        main_entry = ttk.Entry(recipients_frame, textvariable=self.main_email_var,
                               font=("Arial", 10), width=35)
        main_entry.pack(fill=tk.X, pady=(0, 15))

        # Sección de CCs
        cc_label_frame = ttk.Frame(recipients_frame)
        cc_label_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(cc_label_frame, text="📋 Copias (CC):").pack(side=tk.LEFT)
        self.cc_counter_label = ttk.Label(cc_label_frame, text=f"0/{self.max_ccs}",
                                          foreground="gray", font=("Arial", 9))
        self.cc_counter_label.pack(side=tk.RIGHT)

        # Frame con scroll para CCs
        self.create_cc_scroll_area(recipients_frame)

        # Botón agregar CC
        self.add_cc_btn = ttk.Button(recipients_frame, text="➕ Agregar CC",
                                     command=self.add_cc_field)
        self.add_cc_btn.pack(fill=tk.X, pady=(10, 0))

    def create_cc_scroll_area(self, parent):
        """Crea el área con scroll para los campos CC."""
        # Frame con canvas para scroll
        scroll_frame = ttk.Frame(parent)
        scroll_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        scroll_frame.grid_rowconfigure(0, weight=1)
        scroll_frame.grid_columnconfigure(0, weight=1)

        # Canvas y scrollbar
        canvas = tk.Canvas(scroll_frame, bg="white", height=200)
        scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=canvas.yview)

        canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Frame interno para los CCs
        self.cc_frame = ttk.Frame(canvas)
        canvas.create_window((0, 0), window=self.cc_frame, anchor="nw")

        # Configurar scroll
        def configure_scroll_region(event):
            canvas.configure(scrollregion=canvas.bbox("all"))

        self.cc_frame.bind("<Configure>", configure_scroll_region)

        # Scroll con mouse wheel
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def create_control_column(self, parent):
        """Crea la columna de control y estado."""
        control_frame = ttk.LabelFrame(parent, text="🔧 Control y Estado", padding=15)
        control_frame.grid(row=0, column=2, padx=(5, 0), pady=0, sticky="nsew")

        # Estado de Email
        email_status_frame = ttk.LabelFrame(control_frame, text="Estado Email", padding=10)
        email_status_frame.pack(fill=tk.X, pady=(0, 15))

        self.credentials_status_label = ttk.Label(email_status_frame, text="🔴 No configurado",
                                                  font=("Arial", 9), wraplength=180)
        self.credentials_status_label.pack()

        # Estado de Destinatarios
        recipients_status_frame = ttk.LabelFrame(control_frame, text="Estado Destinatarios", padding=10)
        recipients_status_frame.pack(fill=tk.X, pady=(0, 20))

        self.recipients_status_label = ttk.Label(recipients_status_frame, text="🔴 No configurado",
                                                 font=("Arial", 9), wraplength=180)
        self.recipients_status_label.pack()

        # Espaciador
        spacer = ttk.Frame(control_frame)
        spacer.pack(fill=tk.BOTH, expand=True)

        # Botones de acción
        buttons_frame = ttk.Frame(control_frame)
        buttons_frame.pack(fill=tk.X, side=tk.BOTTOM)

        # Botón probar conexión
        test_btn = ttk.Button(buttons_frame, text="🔍 Probar Conexión",
                              command=self.test_connection)
        test_btn.pack(fill=tk.X, pady=(0, 5))

        # Botón guardar todo
        save_btn = ttk.Button(buttons_frame, text="💾 Guardar Todo",
                              command=self.save_all_config)
        save_btn.pack(fill=tk.X, pady=(0, 5))

        # Botón limpiar todo
        clear_btn = ttk.Button(buttons_frame, text="🗑️ Limpiar Todo",
                               command=self.clear_all_config)
        clear_btn.pack(fill=tk.X)

    def add_cc_field(self):
        """Agrega un nuevo campo CC."""
        if len(self.cc_entries) >= self.max_ccs:
            self.update_recipients_status("⚠️ Máximo 10 CCs permitidos", "orange")
            return

        # Frame para este CC
        cc_container = ttk.Frame(self.cc_frame)
        cc_container.pack(fill=tk.X, padx=5, pady=2)
        cc_container.grid_columnconfigure(0, weight=1)

        # Entry para email
        cc_entry = ttk.Entry(cc_container, font=("Arial", 9))
        cc_entry.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        # Botón eliminar
        remove_btn = ttk.Button(cc_container, text="❌", width=3,
                                command=lambda: self.remove_cc_field(cc_container, cc_entry))
        remove_btn.grid(row=0, column=1)

        # Guardar referencia
        self.cc_entries.append({
            'container': cc_container,
            'entry': cc_entry
        })

        self.update_cc_counter()

        # Deshabilitar botón si llegamos al máximo
        if len(self.cc_entries) >= self.max_ccs:
            self.add_cc_btn.config(state="disabled")

    def remove_cc_field(self, container, entry):
        """Elimina un campo CC."""
        # Encontrar y remover de la lista
        self.cc_entries = [cc for cc in self.cc_entries if cc['entry'] != entry]

        # Destruir el container
        container.destroy()

        self.update_cc_counter()

        # Rehabilitar botón agregar
        self.add_cc_btn.config(state="normal")

    def update_cc_counter(self):
        """Actualiza el contador de CCs."""
        if self.cc_counter_label:
            self.cc_counter_label.config(text=f"{len(self.cc_entries)}/{self.max_ccs}")

    def test_connection(self):
        """Prueba la conexión de email con mejor feedback visual."""
        credentials_data = self._get_credentials_data()
        if not all(credentials_data.values()):
            self.update_credentials_status("🔴 Complete todos los campos", "#e74c3c")
            return

        self.update_credentials_status("🔄 Probando conexión...", "#f39c12")

        def test_thread():
            try:
                success, message = self.config_tab.email_manager.test_connection(
                    credentials_data["provider"], credentials_data["email"], credentials_data["password"]
                )

                def update_ui():
                    if success:
                        color = "#27ae60"  # Verde
                        icon = "✅"
                    else:
                        color = "#e74c3c"  # Rojo
                        icon = "❌"
                    self.update_credentials_status(f"{icon} {message}", color)

                self.parent.after(0, update_ui)

            except Exception as e:
                self.parent.after(0, lambda: self.update_credentials_status(f"❌ Error: {str(e)}", "#e74c3c"))

        threading.Thread(target=test_thread, daemon=True).start()

    def save_all_config(self):
        """Guarda toda la configuración con mejor feedback visual."""
        # Validar y obtener datos de credenciales
        credentials_data = self._get_credentials_data()
        if not all(credentials_data.values()):
            self.update_credentials_status("❌ Complete campos de email", "#e74c3c")
            return

        # Validar y obtener datos de destinatarios
        recipients_data = self._get_recipients_data()
        is_valid, error_msg = self._validate_recipients_data(recipients_data)

        if not is_valid:
            self.update_recipients_status(f"❌ Error: {error_msg}", "#e74c3c")
            return

        try:
            # Mostrar estado guardando
            self.update_credentials_status("💾 Guardando...", "#f39c12")
            self.update_recipients_status("💾 Guardando...", "#f39c12")

            # Guardar ambas configuraciones
            combined_config = credentials_data.copy()
            combined_config["recipients_config"] = recipients_data

            self.config_tab.config_manager.save_config(combined_config)

            # Confirmación exitosa
            self.update_credentials_status("✅ Email configurado", "#27ae60")
            self.update_recipients_status("✅ Destinatarios configurados", "#27ae60")

        except Exception as e:
            self.update_credentials_status(f"❌ Error al guardar: {str(e)}", "#e74c3c")
            self.update_recipients_status(f"❌ Error al guardar: {str(e)}", "#e74c3c")

    def clear_all_config(self):
        """Limpia toda la configuración con confirmación visual."""
        # Limpiar credenciales
        self.provider_var.set("Gmail")
        self.email_var.set("")
        self.password_var.set("")

        # Limpiar destinatarios
        self.main_email_var.set("")

        # Eliminar todos los CCs
        for cc_data in self.cc_entries[:]:
            cc_data['container'].destroy()

        self.cc_entries.clear()
        self.update_cc_counter()
        self.add_cc_btn.config(state="normal")

        # Actualizar estados con mejor feedback
        self.update_credentials_status("⚪ No configurado", "#95a5a6")
        self.update_recipients_status("⚪ No configurado", "#95a5a6")

        try:
            self.config_tab.config_manager.clear_config()
        except Exception as e:
            print(f"Error limpiando configuración: {e}")

    def load_existing_config(self):
        """Carga configuración existente con indicadores visuales."""
        try:
            config = self.config_tab.config_manager.load_config()
            if not config:
                self.update_credentials_status("⚪ No configurado", "#95a5a6")
                self.update_recipients_status("⚪ No configurado", "#95a5a6")
                return

            # Cargar credenciales
            self.provider_var.set(config.get("provider", "Gmail"))
            self.email_var.set(config.get("email", ""))
            self.password_var.set(config.get("password", ""))

            if config.get("email"):
                self.update_credentials_status("📋 Email cargado", "#3498db")

            # Cargar destinatarios
            recipients_config = config.get("recipients_config")
            if recipients_config:
                # Cargar destinatario principal
                self.main_email_var.set(recipients_config.get("main_recipient", ""))

                # Cargar CCs
                for cc_email in recipients_config.get("cc_recipients", []):
                    if cc_email.strip():
                        self.add_cc_field()
                        self.cc_entries[-1]['entry'].insert(0, cc_email)

                if recipients_config.get("main_recipient"):
                    self.update_recipients_status("📋 Destinatarios cargados", "#3498db")

        except Exception as e:
            print(f"Error cargando configuración: {e}")
            self.update_credentials_status(f"❌ Error: {str(e)}", "#e74c3c")

    def _get_credentials_data(self):
        """Obtiene los datos de credenciales actuales."""
        return {
            "provider": self.provider_var.get(),
            "email": self.email_var.get().strip(),
            "password": self.password_var.get().strip()
        }

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

    def update_credentials_status(self, message, color):
        """Actualiza el estado de credenciales."""
        if self.credentials_status_label:
            self.credentials_status_label.config(text=message, foreground=color)

    def update_recipients_status(self, message, color):
        """Actualiza el estado de destinatarios."""
        if self.recipients_status_label:
            self.recipients_status_label.config(text=message, foreground=color)

    def show(self):
        """Muestra la sub-pestaña."""
        if not self.is_visible:
            self.is_visible = True

    def hide(self):
        """Oculta la sub-pestaña."""
        if self.is_visible:
            self.is_visible = False