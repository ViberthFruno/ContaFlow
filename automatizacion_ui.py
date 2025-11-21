# automatizacion_ui.py
"""
Interfaz simplificada de automatización usando tkinter nativo.
Controles básicos del bot sin configuración de intervalo ni auto-inicio.
Diseño moderno con log oscuro elegante y controles optimizados.
"""
# Archivos relacionados: automatizacion_tab.py, bot_controller.py, theme_manager.py

import tkinter as tk
from tkinter import ttk
import threading
from datetime import datetime
from theme_manager import ModernTheme, create_modern_text_widget
from email_config_modals import EmailConfigModal, RecipientsConfigModal


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

        # Configurar grid con 3 secciones:
        # - Fila 0: Sección superior grande (columnspan=2)
        # - Fila 1, Col 0: Control del bot (izquierda)
        # - Fila 1, Col 1: Log (derecha)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_columnconfigure(1, weight=3)  # Log más ancho
        self.main_frame.grid_rowconfigure(0, weight=0)  # Sección superior - tamaño fijo
        self.main_frame.grid_rowconfigure(1, weight=1)  # Sección inferior - se expande

    def create_content(self):
        """Crea el contenido con 3 secciones."""
        # Sección superior - Botón principal (fila 0, columna 0-1)
        self.create_top_section()

        # Sección inferior izquierda - Control del bot (fila 1, columna 0)
        self.create_control_panel()

        # Sección inferior derecha - Log (fila 1, columna 1)
        self.create_log_panel()

    def create_top_section(self):
        """Crea la sección superior con el botón principal de iniciar/detener bot."""
        # Frame superior grande que ocupa todo el ancho
        top_frame = ttk.LabelFrame(self.main_frame, text="🤖 Bot de Automatización",
                                   padding=30, style="Modern.TLabelframe")
        top_frame.grid(row=0, column=0, columnspan=2, padx=0, pady=(0, 15), sticky="ew")

        # Configurar para centrar contenido
        top_frame.grid_columnconfigure(0, weight=1)

        # Card con estado del bot
        status_card = tk.Frame(top_frame, bg=ModernTheme.BG_SURFACE,
                              highlightbackground=ModernTheme.BORDER_LIGHT,
                              highlightthickness=1)
        status_card.grid(row=0, column=0, pady=(0, 20))

        # Estado del bot
        self.bot_status_label = tk.Label(status_card, text="🔴 Bot Detenido",
                                         font=("Segoe UI", 16, "bold"),
                                         bg=ModernTheme.BG_SURFACE,
                                         fg=ModernTheme.DANGER,
                                         padx=30, pady=20)
        self.bot_status_label.pack()

        # Botón toggle (iniciar/detener) - Grande y centrado
        self.btn_toggle = ttk.Button(top_frame, text="▶️ Iniciar Bot",
                                     command=self._handle_toggle_bot_click,
                                     style="Primary.TButton")
        self.btn_toggle.grid(row=1, column=0, pady=(0, 10), ipadx=40, ipady=15)

    def create_control_panel(self):
        """Crea el panel de controles con botón de limpiar log."""
        # Frame principal de controles con estilo moderno
        control_frame = ttk.LabelFrame(self.main_frame, text="🎛️ Control del Bot",
                                      padding=20, style="Modern.TLabelframe")
        control_frame.grid(row=1, column=0, padx=(0, 10), pady=0, sticky="nsew")

        # Información del sistema
        self.create_info_section(control_frame)

        # Botones de control
        self.create_control_buttons(control_frame)

    def create_info_section(self, parent):
        """Crea la sección de información del sistema."""
        # Card con información del sistema
        info_card = tk.Frame(parent, bg=ModernTheme.BG_SURFACE,
                            highlightbackground=ModernTheme.BORDER_LIGHT,
                            highlightthickness=1)
        info_card.pack(fill=tk.BOTH, expand=True, pady=(0, 20))

        # Label de título
        title_label = tk.Label(info_card, text="Información del Sistema",
                              font=ModernTheme.FONT_SUBHEADING,
                              bg=ModernTheme.BG_SURFACE,
                              fg=ModernTheme.TEXT_SECONDARY)
        title_label.pack(pady=(15, 10))

        # Información del monitoreo
        info_text = "⏰ Monitoreo: 1 minuto\n🎯 Búsqueda: 'Cargador'\n📎 Archivos: Excel"
        info_label = tk.Label(info_card, text=info_text,
                             font=("Segoe UI", 10),
                             bg=ModernTheme.BG_SURFACE,
                             fg=ModernTheme.TEXT_PRIMARY,
                             justify=tk.LEFT,
                             pady=10)
        info_label.pack()

    def create_control_buttons(self, parent):
        """Crea los botones de control."""
        buttons_frame = ttk.Frame(parent)
        buttons_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(10, 0))

        # Botón configuración de email
        email_config_btn = ttk.Button(buttons_frame, text="⚙️ Configuración de Email",
                                      command=self._handle_email_config_click,
                                      style="TButton")
        email_config_btn.pack(fill=tk.X, pady=(0, 8), ipady=8)

        # Botón configurar destinatarios
        recipients_btn = ttk.Button(buttons_frame, text="📧 Configurar Destinatarios",
                                    command=self._handle_recipients_config_click,
                                    style="TButton")
        recipients_btn.pack(fill=tk.X, pady=(0, 8), ipady=8)

        # Botón limpiar log
        clear_btn = ttk.Button(buttons_frame, text="🗑️ Limpiar Log",
                               command=self._handle_clear_log_click,
                               style="TButton")
        clear_btn.pack(fill=tk.X, ipady=8)

    def create_log_panel(self):
        """Crea el panel de log moderno con fondo oscuro elegante."""
        log_frame = ttk.LabelFrame(self.main_frame, text="📋 Log de Actividad",
                                  padding=10, style="Modern.TLabelframe")
        log_frame.grid(row=1, column=1, padx=0, pady=0, sticky="nsew")

        # Configurar expansión del frame
        log_frame.grid_columnconfigure(0, weight=1)
        log_frame.grid_rowconfigure(0, weight=1)

        # Crear text widget con scrollbar
        text_frame = ttk.Frame(log_frame)
        text_frame.grid(row=0, column=0, sticky="nsew")
        text_frame.grid_columnconfigure(0, weight=1)
        text_frame.grid_rowconfigure(0, weight=1)

        # Text widget moderno con fondo oscuro elegante
        self.log_text = create_modern_text_widget(text_frame,
                                                  wrap=tk.WORD,
                                                  state=tk.DISABLED)
        self.log_text.grid(row=0, column=0, sticky="nsew")

        # Scrollbar moderna
        scrollbar = ttk.Scrollbar(text_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.config(yscrollcommand=scrollbar.set)

        # Mensajes iniciales con estilo
        self.add_log_message("🚀 Sistema de búsqueda 'Cargador' iniciado", "info")
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

    def _handle_email_config_click(self):
        """Maneja clic en configuración de email."""
        try:
            EmailConfigModal(self.parent)
        except Exception as e:
            print(f"Error abriendo modal de email: {e}")
            self.add_log_message(f"❌ Error abriendo configuración de email: {e}", "error")

    def _handle_recipients_config_click(self):
        """Maneja clic en configuración de destinatarios."""
        try:
            RecipientsConfigModal(self.parent)
        except Exception as e:
            print(f"Error abriendo modal de destinatarios: {e}")
            self.add_log_message(f"❌ Error abriendo configuración de destinatarios: {e}", "error")

    # ========== MÉTODOS DE LOGGING ==========

    def add_log_message(self, message, msg_type="info"):
        """Agrega mensaje al log de forma thread-safe."""

        def _add_message_safe():
            try:
                timestamp = datetime.now().strftime("%H:%M:%S")
                formatted_message = f"[{timestamp}] {message}\n"

                if not self.log_text.winfo_exists():
                    return

                # Habilitar edición temporal
                self.log_text.config(state=tk.NORMAL)

                # Agregar mensaje
                self.log_text.insert(tk.END, formatted_message)

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
                self.add_log_message("🗑️ Log limpiado", "info")
                self.add_log_message("🎯 Sistema: Búsqueda 'Cargador' + Excel", "info")
        except Exception as e:
            print(f"Error limpiando log: {e}")

    # ========== MÉTODOS DE ACTUALIZACIÓN DE UI ==========

    def update_bot_status(self, status_text, color):
        """Actualiza el estado visual del bot con colores modernos."""
        try:
            # Mapear colores a la paleta moderna
            color_map = {
                'green': ModernTheme.SUCCESS,
                'red': ModernTheme.DANGER,
                'orange': ModernTheme.WARNING,
                'blue': ModernTheme.INFO
            }
            modern_color = color_map.get(color, color)
            self.bot_status_label.config(text=status_text, fg=modern_color)
        except Exception as e:
            print(f"Error actualizando status: {e}")

    def update_ui_for_running_state(self):
        """Actualiza UI para estado 'corriendo' con estilos modernos."""
        try:
            self.update_bot_status("🟢 Bot Activo - Monitoreo cada 1 min", "green")
            self.btn_toggle.config(text="⏹️ Detener Bot", state=tk.NORMAL, style="Danger.TButton")
        except Exception as e:
            print(f"Error actualizando UI running: {e}")

    def update_ui_for_stopping_state(self):
        """Actualiza UI para estado 'deteniéndose' con estilos modernos."""
        try:
            self.update_bot_status("🟡 Deteniendo...", "orange")
            self.btn_toggle.config(text="⏳ Deteniendo...", state=tk.DISABLED)
        except Exception as e:
            print(f"Error actualizando UI stopping: {e}")

    def update_ui_for_stopped_state(self):
        """Actualiza UI para estado 'detenido' con estilos modernos."""
        try:
            self.update_bot_status("🔴 Bot Detenido", "red")
            self.btn_toggle.config(text="▶️ Iniciar Bot", state=tk.NORMAL, style="Primary.TButton")
        except Exception as e:
            print(f"Error actualizando UI stopped: {e}")

    def reset_stop_state(self, bot_running):
        """Resetea estado de parada con estilos modernos."""
        try:
            if bot_running:
                text = "⏹️ Detener Bot"
                status = "🟢 Bot Activo - Monitoreo cada 1 min"
                color = "green"
                style = "Danger.TButton"
            else:
                text = "▶️ Iniciar Bot"
                status = "🔴 Bot Detenido"
                color = "red"
                style = "Primary.TButton"

            self.btn_toggle.config(text=text, state=tk.NORMAL, style=style)
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