# xml_tab.py
"""
Sub-pestaña de configuración XML con interfaz optimizada usando tkinter nativo.
Configura carpetas base con rutas dinámicas automáticas, actividades comerciales
y opciones de procesamiento Excel/XML con diseño mejorado sin espacios en blanco.
"""
# Archivos relacionados: config_manager.py

import tkinter as tk
from tkinter import ttk, filedialog
import os
from config_manager import ConfigManager


class XmlTab:
    """Sub-pestaña de configuración XML con interfaz gráfica optimizada."""

    def __init__(self, parent):
        """Inicializa la sub-pestaña XML."""
        self.parent = parent
        self.is_visible = False
        self.config_manager = ConfigManager()

        # Definir empresas con sus campos
        self.company_folders = {
            "nargallo": {
                "name": "Nargallo del Este S.A.",
                "folder_var": tk.StringVar(),
                "activity_var": tk.StringVar(),
                "preview_label": None
            },
            "ventas_fruno": {
                "name": "Ventas Fruno, S.A.",
                "folder_var": tk.StringVar(),
                "activity_var": tk.StringVar(),
                "preview_label": None
            },
            "creme_caramel": {
                "name": "Creme Caramel",
                "folder_var": tk.StringVar(),
                "activity_var": tk.StringVar(),
                "preview_label": None
            },
            "su_laka": {
                "name": "Su Laka",
                "folder_var": tk.StringVar(),
                "activity_var": tk.StringVar(),
                "preview_label": None
            }
        }

        # Variables adicionales
        self.output_folder_var = tk.StringVar()
        self.delete_originals_var = tk.BooleanVar(value=True)
        self.auto_send_var = tk.BooleanVar(value=True)
        self.detailed_logs_var = tk.BooleanVar(value=True)
        self.manual_review_limit_var = tk.StringVar(value="3")

        # Referencias a widgets
        self.status_label = None
        self.companies_container = None

        # Crear interfaz
        self.create_interface()
        self.load_xml_config()

    def create_interface(self):
        """Crea la interfaz completa con diseño optimizado."""
        # Frame principal sin grid complejo
        main_frame = ttk.Frame(self.parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Frame superior para mensaje informativo
        self.create_info_section(main_frame)

        # Frame central con dos columnas principales
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 0))

        # Crear las dos columnas
        self.create_folders_section(content_frame)
        self.create_options_section(content_frame)

    def create_info_section(self, parent):
        """Crea la sección informativa superior."""
        info_frame = ttk.LabelFrame(parent, text="💡 Información", padding=8)
        info_frame.pack(fill=tk.X, pady=(0, 10))

        month_info = self.config_manager.get_current_month_folder_info()
        info_text = f"Configure carpetas BASE. Se agregará automáticamente \\{month_info['folder_suffix']} para el procesamiento"

        info_label = ttk.Label(info_frame, text=info_text, foreground="blue",
                               font=("Arial", 9), wraplength=600)
        info_label.pack()

    def create_folders_section(self, parent):
        """Crea la sección de carpetas con scroll optimizado."""
        # Frame izquierdo para carpetas
        left_frame = ttk.LabelFrame(parent, text="🗂️ Configuración de Carpetas XML", padding=10)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))

        # Crear scroll area optimizado
        self.create_optimized_scroll_area(left_frame)

        # Agregar carpeta de salida al final
        self.create_output_section(left_frame)

    def create_optimized_scroll_area(self, parent):
        """Crea área de scroll optimizada sin espacios en blanco excesivos."""
        # Frame contenedor para scroll
        scroll_container = ttk.Frame(parent)
        scroll_container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Canvas con altura fija más pequeña
        canvas = tk.Canvas(scroll_container, highlightthickness=0, height=300)

        # Scrollbar
        scrollbar = ttk.Scrollbar(scroll_container, orient="vertical", command=canvas.yview)

        # Frame scrollable interno
        self.companies_container = ttk.Frame(canvas)

        # Configurar pack para canvas y scrollbar
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Configurar canvas
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas_window = canvas.create_window((0, 0), window=self.companies_container, anchor="nw")

        # Crear secciones de empresas
        for company_key, company_info in self.company_folders.items():
            self.create_company_section(self.companies_container, company_key, company_info)

        # Configurar scroll region dinámico
        def configure_scroll_region(event=None):
            canvas.configure(scrollregion=canvas.bbox("all"))

            # Ajustar ancho del frame interno
            canvas_width = canvas.winfo_width()
            if canvas_width > 1:
                canvas.itemconfig(canvas_window, width=canvas_width)

        self.companies_container.bind("<Configure>", configure_scroll_region)
        canvas.bind("<Configure>", configure_scroll_region)

        # Scroll con rueda del mouse (solo en el canvas)
        def _on_canvas_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

        canvas.bind("<MouseWheel>", _on_canvas_mousewheel)
        canvas.bind("<Enter>", lambda e: canvas.focus_set())

    def create_company_section(self, parent, company_key, company_info):
        """Crea sección compacta para una empresa."""
        company_frame = ttk.LabelFrame(parent, text=f"📁 {company_info['name']}", padding=8)
        company_frame.pack(fill=tk.X, padx=2, pady=3)

        # Carpeta BASE
        ttk.Label(company_frame, text="Carpeta BASE:", font=("Arial", 9)).pack(anchor=tk.W)

        # Frame para entry y botón
        folder_frame = ttk.Frame(company_frame)
        folder_frame.pack(fill=tk.X, pady=(2, 5))

        folder_entry = ttk.Entry(folder_frame, textvariable=company_info['folder_var'],
                                 font=("Arial", 9))
        folder_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        browse_btn = ttk.Button(folder_frame, text="📁", width=3,
                                command=lambda: self.browse_company_folder(company_info))
        browse_btn.pack(side=tk.RIGHT)

        # Preview de ruta dinámica
        preview_label = ttk.Label(company_frame, text="📂 Ingrese carpeta base para ver ruta dinámica",
                                  foreground="gray", font=("Arial", 8), wraplength=350)
        preview_label.pack(anchor=tk.W, pady=(0, 5))
        company_info['preview_label'] = preview_label

        # Actividad comercial
        ttk.Label(company_frame, text="🏢 Actividad comercial (opcional):",
                  font=("Arial", 9)).pack(anchor=tk.W, pady=(5, 2))

        activity_entry = ttk.Entry(company_frame, textvariable=company_info['activity_var'],
                                   font=("Arial", 9))
        activity_entry.pack(fill=tk.X)

        # Vincular eventos para preview
        folder_entry.bind('<KeyRelease>', lambda e: self.update_dynamic_path_preview(company_info))
        folder_entry.bind('<FocusOut>', lambda e: self.update_dynamic_path_preview(company_info))

    def create_output_section(self, parent):
        """Crea sección compacta de carpeta de salida."""
        output_frame = ttk.LabelFrame(parent, text="💾 Carpeta de Archivos Procesados", padding=8)
        output_frame.pack(fill=tk.X, pady=(5, 0))

        folder_frame = ttk.Frame(output_frame)
        folder_frame.pack(fill=tk.X)

        output_entry = ttk.Entry(folder_frame, textvariable=self.output_folder_var, font=("Arial", 9))
        output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        browse_btn = ttk.Button(folder_frame, text="📁", width=3, command=self.browse_output_folder)
        browse_btn.pack(side=tk.RIGHT)

    def create_options_section(self, parent):
        """Crea la sección de opciones."""
        # Frame derecho para opciones
        right_frame = ttk.LabelFrame(parent, text="⚙️ Opciones de Procesamiento", padding=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 0))

        # Opciones de procesamiento
        self.create_processing_options(right_frame)

        # Estado y controles
        self.create_status_and_controls(right_frame)

    def create_processing_options(self, parent):
        """Crea las opciones de procesamiento."""
        options_frame = ttk.LabelFrame(parent, text="Configuración", padding=8)
        options_frame.pack(fill=tk.X, pady=(0, 10))

        # Checkboxes
        ttk.Checkbutton(options_frame, text="Eliminar archivos Excel originales",
                        variable=self.delete_originals_var).pack(anchor=tk.W, pady=1)

        ttk.Checkbutton(options_frame, text="Enviar automáticamente",
                        variable=self.auto_send_var).pack(anchor=tk.W, pady=1)

        ttk.Checkbutton(options_frame, text="Logs detallados",
                        variable=self.detailed_logs_var).pack(anchor=tk.W, pady=1)

        # Límite de revisiones
        ttk.Label(options_frame, text="Máx. detalles antes de revisión manual:",
                  font=("Arial", 9)).pack(anchor=tk.W, pady=(8, 2))

        limit_entry = ttk.Entry(options_frame, textvariable=self.manual_review_limit_var,
                                width=8, font=("Arial", 9))
        limit_entry.pack(anchor=tk.W)

    def create_status_and_controls(self, parent):
        """Crea sección de estado y controles."""
        # Estado
        status_frame = ttk.LabelFrame(parent, text="Estado", padding=8)
        status_frame.pack(fill=tk.X, pady=(0, 15))

        self.status_label = ttk.Label(status_frame, text="🔴 No configurado",
                                      font=("Arial", 9), wraplength=180)
        self.status_label.pack()

        # Botones
        buttons_frame = ttk.Frame(parent)
        buttons_frame.pack(fill=tk.X)

        ttk.Button(buttons_frame, text="💾 Guardar",
                   command=self.save_xml_config).pack(fill=tk.X, pady=(0, 3))

        ttk.Button(buttons_frame, text="🔍 Probar Rutas",
                   command=self.test_dynamic_xml_access).pack(fill=tk.X, pady=(0, 3))

        ttk.Button(buttons_frame, text="🗑️ Limpiar",
                   command=self.clear_xml_config).pack(fill=tk.X)

    def browse_company_folder(self, company_info):
        """Abre diálogo para seleccionar carpeta de empresa."""
        folder_path = filedialog.askdirectory(
            title=f"Seleccionar carpeta BASE para {company_info['name']}",
            initialdir=os.path.expanduser("~")
        )
        if folder_path:
            company_info['folder_var'].set(folder_path)
            self.update_dynamic_path_preview(company_info)

    def browse_output_folder(self):
        """Abre diálogo para seleccionar carpeta de salida."""
        folder_path = filedialog.askdirectory(
            title="Seleccionar carpeta para archivos procesados",
            initialdir=os.path.expanduser("~")
        )
        if folder_path:
            self.output_folder_var.set(folder_path)

    def update_dynamic_path_preview(self, company_info):
        """Actualiza el preview de la ruta dinámica."""
        try:
            if not company_info['preview_label']:
                return

            base_path = company_info['folder_var'].get().strip()

            if base_path:
                dynamic_path = self.config_manager.build_dynamic_xml_path(base_path)
                exists, _, message = self.config_manager.validate_dynamic_xml_path(base_path)

                status_icon = "✅" if exists else "⚠️"
                color = "green" if exists else "orange"

                preview_text = f"📂 {dynamic_path} {status_icon}"
                company_info['preview_label'].config(text=preview_text, foreground=color)
            else:
                company_info['preview_label'].config(
                    text="📂 Ingrese carpeta base para ver ruta dinámica", foreground="gray")

        except Exception as e:
            print(f"Error actualizando preview: {e}")

    def test_dynamic_xml_access(self):
        """Prueba el acceso a carpetas XML con rutas dinámicas."""
        xml_config = self.get_current_xml_config()

        if not xml_config["company_folders"]:
            return self.update_status("🔴 Configure al menos una carpeta", "red")

        dynamic_paths_info = self.config_manager.get_all_dynamic_xml_paths(xml_config["company_folders"])
        month_info = self.config_manager.get_current_month_folder_info()

        existing_paths = 0
        results = []

        for company_key, path_info in dynamic_paths_info.items():
            company_name = self.company_folders[company_key]["name"]

            if path_info['exists']:
                try:
                    xml_count = sum(1 for root, dirs, files in os.walk(path_info['dynamic_path'])
                                    for file in files if file.lower().endswith('.xml'))

                    if xml_count > 0:
                        results.append(f"✅ {company_name}: {xml_count}+ XMLs")
                        existing_paths += 1
                    else:
                        results.append(f"⚠️ {company_name}: Carpeta vacía")
                        existing_paths += 1
                except Exception:
                    results.append(f"❌ {company_name}: Error acceso")
            else:
                results.append(f"📅 {company_name}: Sin {month_info['display_text']}")

        total_companies = len(dynamic_paths_info)
        summary = f"Rutas dinámicas ({month_info['display_text']}):\n{existing_paths}/{total_companies} encontradas\n"
        summary += "\n".join(results[:4])

        if len(results) > 4:
            summary += f"\n..."

        color = "green" if existing_paths == total_companies else "orange" if existing_paths > 0 else "red"
        self.update_status(summary, color)

    def save_xml_config(self):
        """Guarda la configuración XML."""
        xml_config = self.get_current_xml_config()

        if not xml_config["company_folders"]:
            return self.update_status("🔴 Configure al menos una carpeta", "red")

        if not xml_config["output_folder"]:
            return self.update_status("🔴 Especifique carpeta de salida", "red")

        # Validar carpetas BASE
        for company_key, base_folder_path in xml_config["company_folders"].items():
            if not os.path.exists(base_folder_path):
                company_name = self.company_folders[company_key]["name"]
                return self.update_status(f"🔴 Carpeta no existe: {company_name}", "red")

        try:
            os.makedirs(xml_config["output_folder"], exist_ok=True)

            manual_limit = int(xml_config["manual_review_limit"])
            if not 1 <= manual_limit <= 20:
                return self.update_status("🔴 Límite debe estar entre 1 y 20", "red")

            self.config_manager.update_config({"xml_config": xml_config})

            configured_count = len(xml_config["company_folders"])
            activities_count = len([a for a in xml_config.get("commercial_activities", {}).values() if a.strip()])

            month_info = self.config_manager.get_current_month_folder_info()
            dynamic_paths_info = self.config_manager.get_all_dynamic_xml_paths(xml_config["company_folders"])
            existing_dynamic_paths = len([info for info in dynamic_paths_info.values() if info['exists']])

            status_msg = f"🟢 Guardado: {configured_count} carpetas base"
            if activities_count > 0:
                status_msg += f", {activities_count} actividades"
            status_msg += f"\n📅 {month_info['display_text']}: {existing_dynamic_paths}/{configured_count} disponibles"

            self.update_status(status_msg, "green")
            self.refresh_all_previews()

        except ValueError:
            self.update_status("🔴 Límite debe ser un número", "red")
        except Exception as e:
            self.update_status(f"🔴 Error: {str(e)}", "red")

    def load_xml_config(self):
        """Carga la configuración XML existente."""
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
                self.detailed_logs_var.set(xml_config.get("detailed_logs", True))
                self.manual_review_limit_var.set(str(xml_config.get("manual_review_limit", 3)))

                # Actualizar previews
                self.refresh_all_previews()

                configured_count = len([k for k, v in company_folders.items() if v])
                if configured_count > 0:
                    self.update_status(f"🟡 Cargado: {configured_count} carpetas", "orange")
            else:
                # Valor por defecto para carpeta de salida
                default_output = os.path.join(os.path.expanduser("~"), "Downloads", "ContaFlow", "Procesados")
                self.output_folder_var.set(default_output)

        except Exception as e:
            print(f"Error cargando configuración XML: {e}")

    def clear_xml_config(self):
        """Limpia la configuración XML."""
        # Limpiar todas las variables
        for company_info in self.company_folders.values():
            company_info['folder_var'].set("")
            company_info['activity_var'].set("")
            if company_info['preview_label']:
                company_info['preview_label'].config(
                    text="📂 Ingrese carpeta base para ver ruta dinámica", foreground="gray")

        self.output_folder_var.set("")
        self.delete_originals_var.set(True)
        self.auto_send_var.set(True)
        self.detailed_logs_var.set(True)
        self.manual_review_limit_var.set("3")
        self.update_status("🔴 Configuración limpiada", "red")

    def get_current_xml_config(self):
        """Obtiene la configuración XML actual."""
        company_folders = {}
        commercial_activities = {}

        for company_key, company_info in self.company_folders.items():
            folder_path = company_info['folder_var'].get().strip()
            if folder_path:
                company_folders[company_key] = folder_path

            activity = company_info['activity_var'].get().strip()
            commercial_activities[company_key] = activity

        return {
            "company_folders": company_folders,
            "commercial_activities": commercial_activities,
            "output_folder": self.output_folder_var.get().strip(),
            "delete_originals": self.delete_originals_var.get(),
            "auto_send": self.auto_send_var.get(),
            "detailed_logs": self.detailed_logs_var.get(),
            "manual_review_limit": int(
                self.manual_review_limit_var.get()) if self.manual_review_limit_var.get().isdigit() else 3
        }

    def refresh_all_previews(self):
        """Actualiza todos los previews de rutas dinámicas."""
        try:
            for company_info in self.company_folders.values():
                self.update_dynamic_path_preview(company_info)
        except Exception as e:
            print(f"Error refrescando previews: {e}")

    def update_status(self, message, color):
        """Actualiza el estado mostrado."""
        if self.status_label:
            self.status_label.config(text=message, foreground=color)

    def show(self):
        """Muestra la sub-pestaña."""
        if not self.is_visible:
            self.is_visible = True

    def hide(self):
        """Oculta la sub-pestaña."""
        if self.is_visible:
            self.is_visible = False