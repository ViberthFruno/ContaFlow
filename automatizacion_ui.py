# automatizacion_ui.py
"""
Interfaz simplificada de automatización usando tkinter nativo.
Controles básicos del bot sin configuración de intervalo ni auto-inicio.
"""
# Archivos relacionados: automatizacion_tab.py, bot_controller.py

import tkinter as tk
from tkinter import ttk
import threading
from datetime import datetime


class AutomatizacionUI:
    """Interfaz simplificada de automatización con controles básicos únicamente."""

    def __init__(self, parent, config_manager):
        """Inicializa la interfaz simplificada."""
        self.parent = parent
        self.config_manager = config_manager
        self.is_visible = False

        # Variables de UI thread-safe
        self._ui_lock = threading.Lock()

        # Referencias a widgets
        self.main_frame = None
        self.log_text = None
        self.bot_status_label = None
        self.btn_toggle = None

        # Callbacks
        self.on_toggle_bot_click = None
        self.on_clear_log_click = None

    def create_interface(self):
        """Crea la interfaz simplificada."""
        self.create_main_frame()
        self.create_content()

    def create_main_frame(self):
        """Crea el frame principal."""
        self.main_frame = ttk.Frame(self.parent)

        # Configurar grid
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=2)
        self.main_frame.grid_rowconfigure(0, weight=1)

    def create_content(self):
        """Crea el contenido simplificado."""
        # Crear frame izquierdo - Controles básicos
        self.create_control_panel()

        # Crear frame derecho - Log
        self.create_log_panel()

    def create_control_panel(self):
        """Crea el panel de controles simplificado con mejor diseño."""
        # Frame principal de controles
        control_frame = ttk.LabelFrame(self.main_frame, text="🎛️ Control del Bot", padding=15)
        control_frame.grid(row=0, column=0, padx=(0, 10), pady=0, sticky="nsew")

        # Configurar grid para expansión
        control_frame.grid_rowconfigure(1, weight=1)

        # Estado del bot
        self.create_status_section(control_frame)

        # Separador decorativo
        separator = ttk.Separator(control_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=(10, 10))

        # Información del sistema
        self.create_info_section(control_frame)

        # Botones principales
        self.create_main_buttons(control_frame)

    def create_status_section(self, parent):
        """Crea la sección de estado con mejor diseño visual."""
        status_frame = ttk.LabelFrame(parent, text="Estado Actual", padding=15)
        status_frame.pack(fill=tk.X, pady=(0, 10))

        self.bot_status_label = ttk.Label(
            status_frame,
            text="🔴 Bot Detenido",
            font=("Arial", 11, "bold"),
            foreground="#e74c3c"
        )
        self.bot_status_label.pack()

    def create_info_section(self, parent):
        """Crea la sección de información del sistema."""
        info_frame = ttk.LabelFrame(parent, text="Configuración", padding=10)
        info_frame.pack(fill=tk.X, pady=(0, 10))

        # Información del monitoreo
        info_items = [
            ("⏰ Intervalo:", "1 minuto"),
            ("🎯 Objetivo:", "Correo 'Cargador'"),
            ("📄 Archivos:", "Excel (.xlsx/.xls)")
        ]

        for label_text, value_text in info_items:
            item_frame = ttk.Frame(info_frame)
            item_frame.pack(fill=tk.X, pady=2)

            ttk.Label(
                item_frame,
                text=label_text,
                font=("Arial", 9, "bold")
            ).pack(side=tk.LEFT)

            ttk.Label(
                item_frame,
                text=value_text,
                font=("Arial", 9),
                foreground="#7f8c8d"
            ).pack(side=tk.LEFT, padx=(5, 0))

    def create_main_buttons(self, parent):
        """Crea los botones principales con mejor diseño."""
        buttons_frame = ttk.Frame(parent)
        buttons_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(10, 0))

        # Botón toggle (iniciar/detener) más grande y destacado
        self.btn_toggle = ttk.Button(
            buttons_frame,
            text="▶️ Iniciar Bot",
            command=self._handle_toggle_bot_click
        )
        self.btn_toggle.pack(fill=tk.X, pady=(0, 10), ipady=5)

        # Separador
        separator = ttk.Separator(buttons_frame, orient='horizontal')
        separator.pack(fill=tk.X, pady=(5, 10))

        # Botón limpiar log
        clear_btn = ttk.Button(
            buttons_frame,
            text="🗑️ Limpiar Log",
            command=self._handle_clear_log_click
        )
        clear_btn.pack(fill=tk.X)

    def create_log_panel(self):
        """Crea el panel de log de actividad con sistema de colores."""
        log_frame = ttk.LabelFrame(self.main_frame, text="📋 Log de Actividad", padding=10)
        log_frame.grid(row=0, column=1, padx=0, pady=0, sticky="nsew")

        # Configurar expansión del frame
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)

        # Crear text widget con scrollbar
        text_frame = ttk.Frame(log_frame)
        text_frame.grid(row=0, column=0, sticky="nsew")
        text_frame.grid_columnconfigure(0, weight=1)
        text_frame.grid_rowconfigure(0, weight=1)

        # Text widget con mejor configuración
        self.log_text = tk.Text(
            text_frame,
            wrap=tk.WORD,
            state=tk.DISABLED,
            font=("Courier", 9),
            bg="#ffffff",
            fg="#000000",
            relief=tk.SUNKEN,
            borderwidth=2
        )
        self.log_text.grid(row=0, column=0, sticky="nsew")

        # Configurar tags de colores para diferentes niveles de log
        self.log_text.tag_config("error", foreground="#FF0000", font=("Courier", 9, "bold"))      # Rojo
        self.log_text.tag_config("warning", foreground="#FF8C00", font=("Courier", 9, "bold"))    # Naranja
        self.log_text.tag_config("info", foreground="#0066CC", font=("Courier", 9))               # Azul
        self.log_text.tag_config("success", foreground="#008000", font=("Courier", 9, "bold"))    # Verde
        self.log_text.tag_config("debug", foreground="#808080", font=("Courier", 9))              # Gris

        # Scrollbar
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.config(yscrollcommand=scrollbar.set)

        # Mensajes iniciales con colores
        self.add_log_message("🚀 Sistema de búsqueda 'Cargador' iniciado", "success")
        self.add_log_message("⏰ Monitoreo configurado: 1 minuto (fijo)", "info")
        self.add_log_message("🎯 Objetivo: Correos 'Cargador' con archivos Excel", "info")

    # ========== MÉTODOS DE CALLBACKS ==========

    def _handle_toggle_bot_click(self):
        """Maneja clic en botón toggle."""
        if self.on_toggle_bot_click:
            self.on_toggle_bot_click()

    def _handle_clear_log_click(self):
        """Maneja clic en limpiar log."""
        if self.on_clear_log_click:
            self.on_clear_log_click()

    # ========== MÉTODOS DE LOGGING ==========

    def add_log_message(self, message, msg_type="info"):
        """Agrega mensaje al log de forma thread-safe con colores según tipo."""

        def _add_message_safe():
            try:
                timestamp = datetime.now().strftime("%H:%M:%S")
                formatted_message = f"[{timestamp}] {message}\n"

                if not self.log_text.winfo_exists():
                    return

                # Habilitar edición temporal
                self.log_text.config(state=tk.NORMAL)

                # Determinar el tag según el tipo de mensaje
                tag = msg_type if msg_type in ["error", "warning", "info", "success", "debug"] else "info"

                # Agregar mensaje con el color correspondiente
                self.log_text.insert(tk.END, formatted_message, tag)

                # Limitar líneas (mantener últimas 1000)
                content = self.log_text.get("1.0", tk.END)
                lines = content.split('\n')
                if len(lines) > 1000:
                    self.log_text.delete("1.0", tk.END)
                    self.log_text.insert("1.0", '\n'.join(lines[-900:]))

                # Scroll al final
                self.log_text.see(tk.END)

                # Deshabilitar edición
                self.log_text.config(state=tk.DISABLED)

            except Exception as e:
                print(f"Error agregando mensaje al log: {e}")

        # Ejecutar en hilo principal de GUI
        try:
            if self.parent and hasattr(self.parent, 'after'):
                self.parent.after(0, _add_message_safe)
            else:
                _add_message_safe()
        except Exception as e:
            print(f"Error programando mensaje de log: {e}")
            print(f"[{datetime.now().strftime('%H:%M:%S')}] {message}")

    def clear_log(self):
        """Limpia el log de actividad."""
        try:
            if self.log_text.winfo_exists():
                self.log_text.config(state=tk.NORMAL)
                self.log_text.delete("1.0", tk.END)
                self.log_text.config(state=tk.DISABLED)
                self.add_log_message("🗑️ Log limpiado correctamente", "warning")
                self.add_log_message("🎯 Sistema: Búsqueda 'Cargador' + Excel", "info")
        except Exception as e:
            print(f"Error limpiando log: {e}")

    # ========== MÉTODOS DE ACTUALIZACIÓN DE UI ==========

    def update_bot_status(self, status_text, color):
        """Actualiza el estado visual del bot con colores mejorados."""
        try:
            # Mapeo de colores a hex para mejor visualización
            color_map = {
                "green": "#27ae60",
                "red": "#e74c3c",
                "orange": "#f39c12",
                "blue": "#3498db"
            }
            hex_color = color_map.get(color, color)
            self.bot_status_label.config(text=status_text, foreground=hex_color)
        except Exception as e:
            print(f"Error actualizando status: {e}")

    def update_ui_for_running_state(self):
        """Actualiza UI para estado 'corriendo' con diseño mejorado."""
        try:
            self.update_bot_status("🟢 Bot Activo - Monitoreo Activo", "green")
            self.btn_toggle.config(text="⏹️ Detener Bot", state=tk.NORMAL)
            self.add_log_message("✅ Bot iniciado correctamente", "success")
        except Exception as e:
            print(f"Error actualizando UI running: {e}")

    def update_ui_for_stopping_state(self):
        """Actualiza UI para estado 'deteniéndose' con mejor feedback."""
        try:
            self.update_bot_status("🟡 Deteniendo Sistema...", "orange")
            self.btn_toggle.config(text="⏳ Deteniendo...", state=tk.DISABLED)
            self.add_log_message("⏸️ Deteniendo bot...", "warning")
        except Exception as e:
            print(f"Error actualizando UI stopping: {e}")

    def update_ui_for_stopped_state(self):
        """Actualiza UI para estado 'detenido' con confirmación."""
        try:
            self.update_bot_status("🔴 Bot Detenido", "red")
            self.btn_toggle.config(text="▶️ Iniciar Bot", state=tk.NORMAL)
            self.add_log_message("⏹️ Bot detenido", "warning")
        except Exception as e:
            print(f"Error actualizando UI stopped: {e}")

    def reset_stop_state(self, bot_running):
        """Resetea estado de parada con actualización visual completa."""
        try:
            if bot_running:
                text = "⏹️ Detener Bot"
                status = "🟢 Bot Activo - Monitoreo Activo"
                color = "green"
            else:
                text = "▶️ Iniciar Bot"
                status = "🔴 Bot Detenido"
                color = "red"

            self.btn_toggle.config(text=text, state=tk.NORMAL)
            self.update_bot_status(status, color)
        except Exception as e:
            print(f"Error reseteando estado: {e}")

    # ========== MÉTODOS DE CONFIGURACIÓN ==========

    def set_callbacks(self, toggle_bot_callback, clear_log_callback):
        """Configura callbacks desde el controller (simplificado)."""
        self.on_toggle_bot_click = toggle_bot_callback
        self.on_clear_log_click = clear_log_callback

    # ========== MÉTODOS DE VISIBILIDAD ==========

    def show(self):
        """Muestra la interfaz."""
        if not self.is_visible:
            try:
                self.main_frame.pack(fill=tk.BOTH, expand=True)
                self.is_visible = True
            except Exception as e:
                print(f"Error mostrando interfaz: {e}")

    def hide(self):
        """Oculta la interfaz."""
        if self.is_visible:
            try:
                self.main_frame.pack_forget()
                self.is_visible = False
            except Exception as e:
                print(f"Error ocultando interfaz: {e}")

    # ========== MÉTODOS DE COMPATIBILIDAD (SIMPLIFICADOS) ==========

    def update_statistics(self, emails_found=None, files_downloaded=None):
        """Método mantenido para compatibilidad (simplificado)."""
        if emails_found is not None or files_downloaded is not None:
            stats_msg = f"📊 Estadísticas - Emails: {emails_found or 0}, Archivos: {files_downloaded or 0}"
            self.add_log_message(stats_msg, "info")