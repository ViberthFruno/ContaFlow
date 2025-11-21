# configuracion_tab.py
"""
Pestaña simplificada de configuración usando tkinter nativo.
Contiene sub-pestañas para email/destinatarios, búsqueda y XML únicamente.
Diseño moderno con cards y estilos elegantes.
"""
# Archivos relacionados: busqueda_tab.py, xml_tab.py, config_manager.py, email_manager.py, theme_manager.py

import tkinter as tk
from tkinter import ttk
import threading
import re
from email_manager import EmailManager
from config_manager import ConfigManager
from combustible_exclusions_tab import CombustibleExclusionsTab
from theme_manager import ModernTheme


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
        """Crea el título moderno de la pestaña."""
        title_frame = tk.Frame(self.main_frame, bg=ModernTheme.BG_SURFACE,
                              highlightbackground=ModernTheme.BORDER_LIGHT,
                              highlightthickness=1)
        title_frame.grid(row=0, column=0, padx=10, pady=(10, 5), sticky="ew")

        title_label = tk.Label(title_frame, text="⚙️ Configuración",
                              font=ModernTheme.FONT_HEADING,
                              bg=ModernTheme.BG_SURFACE,
                              fg=ModernTheme.PRIMARY)
        title_label.pack(pady=(10, 5))

        subtitle_label = tk.Label(title_frame, text="Gestiona todas las configuraciones del bot",
                                 font=ModernTheme.FONT_NORMAL,
                                 bg=ModernTheme.BG_SURFACE,
                                 fg=ModernTheme.TEXT_SECONDARY)
        subtitle_label.pack(pady=(0, 10))

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
        self.next_cc_id = 0  # ✅ Contador para IDs únicos
        self.config_loaded = False  # ✅ Flag para evitar recargas múltiples
        self.is_processing = False  # ✅ Flag para debouncing

        # Referencias a widgets
        self.credentials_status_label = None
        self.recipients_status_label = None
        self.cc_frame = None
        self.cc_counter_label = None
        self.add_cc_btn = None
        self.canvas = None  # ✅ Referencia al canvas para limpiar bindings
        self.mousewheel_binding = None  # ✅ Guardar binding para limpiarlo después

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
        """Crea la columna moderna de configuración de email."""
        email_frame = ttk.LabelFrame(parent, text="📧 Configuración de Email",
                                    padding=15, style="Modern.TLabelframe")
        email_frame.grid(row=0, column=0, padx=(0, 5), pady=0, sticky="nsew")

        # Proveedor
        ttk.Label(email_frame, text="Proveedor de Correo:",
                 font=ModernTheme.FONT_NORMAL).pack(anchor=tk.W, pady=(0, 5))
        provider_combo = ttk.Combobox(email_frame, textvariable=self.provider_var,
                                      values=["Gmail", "Outlook", "Yahoo", "Otro"],
                                      state="readonly", width=30, font=ModernTheme.FONT_NORMAL)
        provider_combo.pack(fill=tk.X, pady=(0, 15))

        # Email
        ttk.Label(email_frame, text="Correo Electrónico:",
                 font=ModernTheme.FONT_NORMAL).pack(anchor=tk.W, pady=(0, 5))
        email_entry = ttk.Entry(email_frame, textvariable=self.email_var,
                               width=35, font=ModernTheme.FONT_NORMAL)
        email_entry.pack(fill=tk.X, pady=(0, 15))

        # Contraseña
        ttk.Label(email_frame, text="Contraseña:",
                 font=ModernTheme.FONT_NORMAL).pack(anchor=tk.W, pady=(0, 5))
        password_entry = ttk.Entry(email_frame, textvariable=self.password_var,
                                   show="*", width=35, font=ModernTheme.FONT_NORMAL)
        password_entry.pack(fill=tk.X, pady=(0, 15))

        # Nota informativa moderna
        note_frame = tk.Frame(email_frame, bg=ModernTheme.INFO,
                             highlightbackground=ModernTheme.SECONDARY,
                             highlightthickness=1)
        note_frame.pack(fill=tk.X, pady=(10, 0))

        note_label = tk.Label(note_frame, text="💡 Para Gmail usa una\ncontraseña de aplicación",
                             fg=ModernTheme.TEXT_WHITE, bg=ModernTheme.INFO,
                             font=ModernTheme.FONT_SMALL, justify=tk.CENTER, pady=8)
        note_label.pack()

        # Espaciador para empujar contenido hacia arriba
        spacer = ttk.Frame(email_frame)
        spacer.pack(fill=tk.BOTH, expand=True)

    def create_recipients_column(self, parent):
        """Crea la columna moderna de configuración de destinatarios."""
        recipients_frame = ttk.LabelFrame(parent, text="👥 Configuración de Destinatarios",
                                         padding=15, style="Modern.TLabelframe")
        recipients_frame.grid(row=0, column=1, padx=5, pady=0, sticky="nsew")
        recipients_frame.grid_rowconfigure(2, weight=1)

        # Destinatario principal
        ttk.Label(recipients_frame, text="📧 Destinatario Principal:",
                 font=ModernTheme.FONT_NORMAL).pack(anchor=tk.W, pady=(0, 5))
        main_entry = ttk.Entry(recipients_frame, textvariable=self.main_email_var,
                               font=ModernTheme.FONT_NORMAL, width=35)
        main_entry.pack(fill=tk.X, pady=(0, 15))

        # ✅ Agregar validación en tiempo real al destinatario principal
        main_entry.bind('<KeyRelease>', self._validate_duplicates_realtime)
        main_entry.bind('<FocusOut>', self._validate_email_format_realtime)

        # Sección de CCs
        cc_label_frame = ttk.Frame(recipients_frame)
        cc_label_frame.pack(fill=tk.X, pady=(0, 5))

        ttk.Label(cc_label_frame, text="📋 Copias (CC):",
                 font=ModernTheme.FONT_NORMAL).pack(side=tk.LEFT)
        self.cc_counter_label = ttk.Label(cc_label_frame, text=f"0/{self.max_ccs}",
                                          foreground=ModernTheme.TEXT_SECONDARY,
                                          font=ModernTheme.FONT_SMALL)
        self.cc_counter_label.pack(side=tk.RIGHT)

        # Frame con scroll para CCs
        self.create_cc_scroll_area(recipients_frame)

        # Botón agregar CC con estilo
        self.add_cc_btn = ttk.Button(recipients_frame, text="➕ Agregar CC",
                                     command=self.add_cc_field,
                                     style="TButton")
        self.add_cc_btn.pack(fill=tk.X, pady=(10, 0), ipady=4)

    def create_cc_scroll_area(self, parent):
        """Crea el área con scroll para los campos CC."""
        # Frame con canvas para scroll
        scroll_frame = ttk.Frame(parent)
        scroll_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 5))
        scroll_frame.grid_rowconfigure(0, weight=1)
        scroll_frame.grid_columnconfigure(0, weight=1)

        # Canvas y scrollbar
        self.canvas = tk.Canvas(scroll_frame, bg="white", height=200)
        scrollbar = ttk.Scrollbar(scroll_frame, orient="vertical", command=self.canvas.yview)

        self.canvas.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.canvas.configure(yscrollcommand=scrollbar.set)

        # Frame interno para los CCs
        self.cc_frame = ttk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.cc_frame, anchor="nw")

        # Configurar scroll
        def configure_scroll_region(event):
            self.canvas.configure(scrollregion=self.canvas.bbox("all"))

        self.cc_frame.bind("<Configure>", configure_scroll_region)

        # ✅ FIX: Usar binding local en lugar de bind_all para evitar memory leaks
        def _on_mousewheel(event):
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        # Binding solo cuando el mouse está sobre el canvas
        def _bind_mousewheel(event):
            self.canvas.bind("<MouseWheel>", _on_mousewheel)

        def _unbind_mousewheel(event):
            self.canvas.unbind("<MouseWheel>")

        self.canvas.bind("<Enter>", _bind_mousewheel)
        self.canvas.bind("<Leave>", _unbind_mousewheel)

    def create_control_column(self, parent):
        """Crea la columna moderna de control y estado."""
        control_frame = ttk.LabelFrame(parent, text="🔧 Control y Estado",
                                      padding=15, style="Modern.TLabelframe")
        control_frame.grid(row=0, column=2, padx=(5, 0), pady=0, sticky="nsew")

        # Estado de Email - Card moderno
        email_status_card = tk.Frame(control_frame, bg=ModernTheme.BG_SURFACE,
                                    highlightbackground=ModernTheme.BORDER_LIGHT,
                                    highlightthickness=1)
        email_status_card.pack(fill=tk.X, pady=(0, 15))

        tk.Label(email_status_card, text="Estado Email",
                font=ModernTheme.FONT_SUBHEADING,
                bg=ModernTheme.BG_SURFACE,
                fg=ModernTheme.TEXT_PRIMARY).pack(pady=(8, 5))

        self.credentials_status_label = tk.Label(email_status_card, text="🔴 No configurado",
                                                font=ModernTheme.FONT_SMALL,
                                                bg=ModernTheme.BG_SURFACE,
                                                fg=ModernTheme.TEXT_PRIMARY,
                                                wraplength=180)
        self.credentials_status_label.pack(pady=(0, 8))

        # Estado de Destinatarios - Card moderno
        recipients_status_card = tk.Frame(control_frame, bg=ModernTheme.BG_SURFACE,
                                         highlightbackground=ModernTheme.BORDER_LIGHT,
                                         highlightthickness=1)
        recipients_status_card.pack(fill=tk.X, pady=(0, 20))

        tk.Label(recipients_status_card, text="Estado Destinatarios",
                font=ModernTheme.FONT_SUBHEADING,
                bg=ModernTheme.BG_SURFACE,
                fg=ModernTheme.TEXT_PRIMARY).pack(pady=(8, 5))

        self.recipients_status_label = tk.Label(recipients_status_card, text="🔴 No configurado",
                                               font=ModernTheme.FONT_SMALL,
                                               bg=ModernTheme.BG_SURFACE,
                                               fg=ModernTheme.TEXT_PRIMARY,
                                               wraplength=180)
        self.recipients_status_label.pack(pady=(0, 8))

        # Espaciador
        spacer = ttk.Frame(control_frame)
        spacer.pack(fill=tk.BOTH, expand=True)

        # Botones de acción con estilos modernos
        buttons_frame = ttk.Frame(control_frame)
        buttons_frame.pack(fill=tk.X, side=tk.BOTTOM)

        # Botón probar conexión
        test_btn = ttk.Button(buttons_frame, text="🔍 Probar Conexión",
                              command=self.test_connection,
                              style="TButton")
        test_btn.pack(fill=tk.X, pady=(0, 8), ipady=5)

        # Botón guardar todo (destacado)
        save_btn = ttk.Button(buttons_frame, text="💾 Guardar Todo",
                              command=self.save_all_config,
                              style="Success.TButton")
        save_btn.pack(fill=tk.X, pady=(0, 8), ipady=5)

        # Botón limpiar todo
        clear_btn = ttk.Button(buttons_frame, text="🗑️ Limpiar Todo",
                               command=self.clear_all_config,
                               style="TButton")
        clear_btn.pack(fill=tk.X, ipady=5)

    def add_cc_field(self):
        """Agrega un nuevo campo CC con ID único y validación en tiempo real."""
        # ✅ Debouncing: prevenir clicks múltiples
        if self.is_processing:
            return

        if len(self.cc_entries) >= self.max_ccs:
            self.update_recipients_status("⚠️ Máximo 10 CCs permitidos", "orange")
            return

        # ✅ Activar flag de procesamiento
        self.is_processing = True
        self.add_cc_btn.config(state="disabled")

        # ✅ Generar ID único para este CC
        cc_id = self.next_cc_id
        self.next_cc_id += 1

        # Frame para este CC
        cc_container = ttk.Frame(self.cc_frame)
        cc_container.pack(fill=tk.X, padx=5, pady=2)
        cc_container.grid_columnconfigure(0, weight=1)

        # Entry para email
        cc_entry = ttk.Entry(cc_container, font=("Arial", 9))
        cc_entry.grid(row=0, column=0, sticky="ew", padx=(0, 5))

        # ✅ Agregar validación en tiempo real
        cc_entry.bind('<KeyRelease>', self._validate_duplicates_realtime)
        cc_entry.bind('<FocusOut>', self._validate_email_format_realtime)

        # Botón eliminar con ID único
        remove_btn = ttk.Button(cc_container, text="❌", width=3,
                                command=lambda: self.remove_cc_field(cc_id))
        remove_btn.grid(row=0, column=1)

        # ✅ Guardar referencia con ID único
        self.cc_entries.append({
            'id': cc_id,
            'container': cc_container,
            'entry': cc_entry
        })

        self.update_cc_counter()

        # ✅ Rehabilitar botón después de 300ms (debouncing)
        def enable_button():
            self.is_processing = False
            if len(self.cc_entries) < self.max_ccs:
                self.add_cc_btn.config(state="normal")

        self.parent.after(300, enable_button)

        # ✅ Feedback visual
        self.update_recipients_status("✅ CC agregado", "green")
        self.parent.after(2000, lambda: self.update_recipients_status("", "green"))

    def remove_cc_field(self, cc_id):
        """Elimina un campo CC usando ID único."""
        # ✅ Buscar CC por ID (más confiable que comparar objetos)
        cc_to_remove = None
        for cc in self.cc_entries:
            if cc['id'] == cc_id:
                cc_to_remove = cc
                break

        if not cc_to_remove:
            return

        # Destruir el container
        cc_to_remove['container'].destroy()

        # ✅ Remover de la lista por ID
        self.cc_entries = [cc for cc in self.cc_entries if cc['id'] != cc_id]

        self.update_cc_counter()

        # Rehabilitar botón agregar
        if not self.is_processing:
            self.add_cc_btn.config(state="normal")

        # ✅ Revalidar duplicados después de eliminar
        self._validate_duplicates_realtime()

        # ✅ Feedback visual
        self.update_recipients_status("🗑️ CC eliminado", "orange")
        self.parent.after(2000, lambda: self.update_recipients_status("", "green"))

    def update_cc_counter(self):
        """Actualiza el contador de CCs."""
        if self.cc_counter_label:
            self.cc_counter_label.config(text=f"{len(self.cc_entries)}/{self.max_ccs}")

    def test_connection(self):
        """Prueba la conexión de email."""
        credentials_data = self._get_credentials_data()
        if not all(credentials_data.values()):
            return self.update_credentials_status("🔴 Complete todos los campos", "red")

        self.update_credentials_status("🔄 Probando conexión...", "orange")

        def test_thread():
            try:
                success, message = self.config_tab.email_manager.test_connection(
                    credentials_data["provider"], credentials_data["email"], credentials_data["password"]
                )

                def update_ui():
                    color = "green" if success else "red"
                    icon = "🟢" if success else "🔴"
                    self.update_credentials_status(f"{icon} {message}", color)

                self.parent.after(0, update_ui)

            except Exception as e:
                self.parent.after(0, lambda: self.update_credentials_status(f"🔴 Error: {str(e)}", "red"))

        threading.Thread(target=test_thread, daemon=True).start()

    def save_all_config(self):
        """Guarda toda la configuración (credenciales + destinatarios)."""
        # Validar y obtener datos de credenciales
        credentials_data = self._get_credentials_data()
        if not all(credentials_data.values()):
            return self.update_credentials_status("🔴 Complete campos de email", "red")

        # Validar y obtener datos de destinatarios
        recipients_data = self._get_recipients_data()
        is_valid, error_msg = self._validate_recipients_data(recipients_data)

        if not is_valid:
            return self.update_recipients_status(f"🔴 Error: {error_msg}", "red")

        try:
            # Guardar ambas configuraciones
            combined_config = credentials_data.copy()
            combined_config["recipients_config"] = recipients_data

            self.config_tab.config_manager.save_config(combined_config)

            self.update_credentials_status("🟢 Email configurado", "green")
            self.update_recipients_status("🟢 Destinatarios configurados", "green")

        except Exception as e:
            self.update_credentials_status(f"🔴 Error al guardar: {str(e)}", "red")
            self.update_recipients_status(f"🔴 Error al guardar: {str(e)}", "red")

    def clear_all_config(self):
        """Limpia toda la configuración."""
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
        self.next_cc_id = 0  # ✅ Resetear contador de IDs
        self.update_cc_counter()
        self.add_cc_btn.config(state="normal")

        # ✅ Resetear flag de configuración cargada
        self.config_loaded = False

        # Actualizar estados
        self.update_credentials_status("🔴 No configurado", "red")
        self.update_recipients_status("🔴 No configurado", "red")

        try:
            self.config_tab.config_manager.clear_config()
        except Exception as e:
            print(f"Error limpiando configuración: {e}")

    def load_existing_config(self):
        """Carga configuración existente."""
        try:
            # ✅ FIX CRÍTICO: Prevenir cargas múltiples
            if self.config_loaded:
                return

            config = self.config_tab.config_manager.load_config()
            if not config:
                return

            # ✅ FIX CRÍTICO: Limpiar CCs existentes ANTES de cargar
            # Esto previene duplicados al recargar
            for cc_data in self.cc_entries[:]:
                cc_data['container'].destroy()
            self.cc_entries.clear()
            self.next_cc_id = 0  # Resetear contador de IDs
            self.update_cc_counter()

            # Cargar credenciales
            self.provider_var.set(config.get("provider", "Gmail"))
            self.email_var.set(config.get("email", ""))
            self.password_var.set(config.get("password", ""))

            if config.get("email"):
                self.update_credentials_status("🟡 Email cargado", "orange")

            # Cargar destinatarios
            recipients_config = config.get("recipients_config")
            if recipients_config:
                # Cargar destinatario principal
                self.main_email_var.set(recipients_config.get("main_recipient", ""))

                # Cargar CCs
                for cc_email in recipients_config.get("cc_recipients", []):
                    if cc_email.strip():
                        # Deshabilitar temporalmente is_processing para cargar
                        original_processing = self.is_processing
                        self.is_processing = False

                        self.add_cc_field()
                        self.cc_entries[-1]['entry'].insert(0, cc_email)

                        self.is_processing = original_processing

                if recipients_config.get("main_recipient"):
                    self.update_recipients_status("🟡 Destinatarios cargados", "orange")

            # ✅ Marcar como cargado
            self.config_loaded = True

        except Exception as e:
            print(f"Error cargando configuración: {e}")

    def _validate_duplicates_realtime(self, event=None):
        """Valida duplicados en tiempo real mientras el usuario escribe."""
        try:
            # Obtener todos los emails actuales
            main_email = self.main_email_var.get().strip().lower()
            cc_emails = []

            for cc_data in self.cc_entries:
                email = cc_data['entry'].get().strip().lower()
                cc_emails.append((cc_data, email))

            # Encontrar duplicados
            all_emails = [main_email] if main_email else []
            seen_emails = set()
            duplicates = set()

            # Agregar main_email a seen
            if main_email:
                seen_emails.add(main_email)

            # Buscar duplicados en CCs
            for cc_data, email in cc_emails:
                if email:
                    if email in seen_emails:
                        duplicates.add(email)
                    seen_emails.add(email)

            # ✅ Marcar visualmente los duplicados
            for cc_data, email in cc_emails:
                entry = cc_data['entry']
                if email and email in duplicates:
                    # Email duplicado - marcar en rojo
                    entry.config(foreground='red')
                elif email and not self._validate_email_format(email):
                    # Formato inválido - marcar en naranja
                    entry.config(foreground='orange')
                else:
                    # Email válido - color normal
                    entry.config(foreground='black')

            # Actualizar estado
            if duplicates:
                dup_list = ', '.join(duplicates)
                self.update_recipients_status(f"⚠️ Duplicados: {dup_list[:30]}...", "orange")
            else:
                # No mostrar nada si no hay errores
                pass

        except Exception as e:
            print(f"Error validando duplicados: {e}")

    def _validate_email_format_realtime(self, event=None):
        """Valida formato de email en tiempo real al perder el foco."""
        try:
            widget = event.widget if event else None
            if not widget:
                return

            email = widget.get().strip()
            if not email:
                widget.config(foreground='black')
                return

            if self._validate_email_format(email):
                # Formato válido - revisar si hay duplicados
                self._validate_duplicates_realtime()
            else:
                # Formato inválido
                widget.config(foreground='orange')
                self.update_recipients_status(f"⚠️ Formato de email inválido", "orange")
                self.parent.after(3000, lambda: self.update_recipients_status("", "green"))

        except Exception as e:
            print(f"Error validando formato: {e}")

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
        """Actualiza el estado de credenciales con colores modernos."""
        if self.credentials_status_label:
            # Mapear colores a la paleta moderna
            color_map = {
                'green': ModernTheme.SUCCESS,
                'red': ModernTheme.DANGER,
                'orange': ModernTheme.WARNING,
                'blue': ModernTheme.INFO
            }
            modern_color = color_map.get(color, color)
            self.credentials_status_label.config(text=message, fg=modern_color)

    def update_recipients_status(self, message, color):
        """Actualiza el estado de destinatarios con colores modernos."""
        if self.recipients_status_label:
            # Mapear colores a la paleta moderna
            color_map = {
                'green': ModernTheme.SUCCESS,
                'red': ModernTheme.DANGER,
                'orange': ModernTheme.WARNING,
                'blue': ModernTheme.INFO
            }
            modern_color = color_map.get(color, color)
            self.recipients_status_label.config(text=message, fg=modern_color)

    def show(self):
        """Muestra la sub-pestaña."""
        if not self.is_visible:
            self.is_visible = True

    def hide(self):
        """Oculta la sub-pestaña."""
        if self.is_visible:
            self.is_visible = False