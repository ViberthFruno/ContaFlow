# main_window.py
"""
Ventana principal simplificada de ContaFlow usando tkinter nativo.
Interfaz limpia con navegación por pestañas sin auto-inicio del bot.
Diseño moderno y optimizado con theme_manager.
"""
# Archivos relacionados: automatizacion_tab.py, configuracion_tab.py, theme_manager.py

import tkinter as tk
from tkinter import ttk
import sys
import threading
import time
from contaflow.ui.theme_manager import ModernTheme, apply_modern_theme


class MainWindow(tk.Tk):
    """Ventana principal simplificada con tkinter nativo sin auto-inicio."""

    def __init__(self):
        """Inicializa la ventana principal simplificada con diseño moderno."""
        super().__init__()
        print("🏗️ Inicializando ventana principal de ContaFlow v2.0 con diseño moderno...")

        # Variables de control simplificadas
        self.tabs = {}
        self.current_tab = None
        self.status_label = None

        # Aplicar tema moderno (primero para mejor rendimiento)
        apply_modern_theme(self)

        # Configurar la ventana
        self.setup_window()

        # Crear interfaz
        self.create_interface()

        # Inicializar pestañas
        self.initialize_tabs()

        # Mostrar pestaña por defecto
        self.show_tab("automatizacion")

        # Actualizar barra de estado
        self.update_status("🟢 Sistema listo", "success")

        print("✅ ContaFlow v2.0 - Sistema Simplificado iniciado correctamente con diseño moderno")

    def setup_window(self):
        """Configura las propiedades básicas de la ventana."""
        try:
            # Título y dimensiones
            self.title("Bot ContaFlow")
            self.geometry("1200x800")
            self.minsize(800, 500)

            # Centrar ventana
            self.center_window()

            # Configurar cierre
            self.protocol("WM_DELETE_WINDOW", self.on_closing)

            # Configurar icono (opcional)
            try:
                self.iconbitmap("icon.ico")
            except:
                pass

        except Exception as e:
            print(f"⚠️ Error configurando ventana: {e}")

    def center_window(self):
        """Centra la ventana en la pantalla."""
        try:
            self.update_idletasks()
            width = self.winfo_width()
            height = self.winfo_height()
            x = (self.winfo_screenwidth() // 2) - (width // 2)
            y = (self.winfo_screenheight() // 2) - (height // 2)
            self.geometry(f"{width}x{height}+{x}+{y}")
        except Exception as e:
            print(f"⚠️ Error centrando ventana: {e}")

    def create_interface(self):
        """Crea la interfaz principal moderna con notebook de pestañas."""
        # Frame principal
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Header con título moderno
        self.create_header(main_frame)

        # Notebook para pestañas (con estilo moderno)
        notebook_container = ttk.Frame(main_frame)
        notebook_container.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 5))

        self.notebook = ttk.Notebook(notebook_container)
        self.notebook.pack(fill=tk.BOTH, expand=True)

        # Crear frames para cada pestaña
        self.automatizacion_frame = ttk.Frame(self.notebook)
        self.configuracion_frame = ttk.Frame(self.notebook)

        # Agregar pestañas al notebook con estilo moderno
        self.notebook.add(self.automatizacion_frame, text="⚡ Automatización")
        self.notebook.add(self.configuracion_frame, text="⚙️ Configuración")

        # Vincular evento de cambio de pestaña
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        # Barra de estado moderna
        self.create_status_bar(main_frame)

    def create_header(self, parent):
        """Crea el header moderno con título."""
        header_frame = tk.Frame(parent, bg=ModernTheme.PRIMARY, height=80)
        header_frame.pack(fill=tk.X, padx=0, pady=0)
        header_frame.pack_propagate(False)

        # Título principal
        title_label = tk.Label(header_frame,
                              text="🤖 Bot ContaFlow",
                              font=ModernTheme.FONT_TITLE,
                              bg=ModernTheme.PRIMARY,
                              fg=ModernTheme.TEXT_WHITE)
        title_label.pack(side=tk.LEFT, padx=20, pady=20)

        # Versión
        version_label = tk.Label(header_frame,
                                text="v2.0",
                                font=ModernTheme.FONT_SMALL,
                                bg=ModernTheme.PRIMARY,
                                fg=ModernTheme.TEXT_LIGHT)
        version_label.pack(side=tk.LEFT, padx=0, pady=25)

    def create_status_bar(self, parent):
        """Crea la barra de estado moderna."""
        status_frame = tk.Frame(parent, bg=ModernTheme.PRIMARY_LIGHT, height=30)
        status_frame.pack(fill=tk.X, side=tk.BOTTOM)
        status_frame.pack_propagate(False)

        # Label de estado
        self.status_label = tk.Label(status_frame,
                                     text="🟢 Sistema listo",
                                     font=ModernTheme.FONT_SMALL,
                                     bg=ModernTheme.PRIMARY_LIGHT,
                                     fg=ModernTheme.TEXT_LIGHT,
                                     anchor=tk.W)
        self.status_label.pack(side=tk.LEFT, padx=15, fill=tk.X, expand=True)

    def update_status(self, message, status_type="info"):
        """
        Actualiza la barra de estado de forma optimizada.

        Args:
            message: Mensaje a mostrar
            status_type: 'success', 'warning', 'danger', 'info'
        """
        if not self.status_label:
            return

        icons = {
            'success': '🟢',
            'warning': '🟡',
            'danger': '🔴',
            'info': '🔵'
        }

        icon = icons.get(status_type, '🔵')
        self.status_label.config(text=f"{icon} {message}")

    def initialize_tabs(self):
        """Inicializa las pestañas del sistema simplificado."""
        try:
            # Importar y crear pestaña de automatización
            from contaflow.ui.tabs.automatizacion_tab import AutomatizacionTab
            self.tabs["automatizacion"] = AutomatizacionTab(self.automatizacion_frame)
            print("✅ Pestaña de automatización inicializada - Sistema simplificado")
        except Exception as e:
            print(f"⚠️ Error inicializando automatización: {e}")
            self.tabs["automatizacion"] = None

        try:
            # Importar y crear pestaña de configuración
            from contaflow.ui.tabs.configuracion_tab import ConfiguracionTab
            self.tabs["configuracion"] = ConfiguracionTab(self.configuracion_frame)
            print("✅ Pestaña de configuración inicializada")
        except Exception as e:
            print(f"⚠️ Error inicializando configuración: {e}")
            self.tabs["configuracion"] = None

    def _on_tab_changed(self, event):
        """Maneja el cambio de pestaña."""
        try:
            selected_tab = event.widget.tab('current')['text']

            if "Automatización" in selected_tab:
                self.show_tab("automatizacion")
            elif "Configuración" in selected_tab:
                self.show_tab("configuracion")

        except Exception as e:
            print(f"⚠️ Error en cambio de pestaña: {e}")

    def show_tab(self, tab_name):
        """Muestra la pestaña especificada."""
        try:
            if tab_name not in self.tabs or self.tabs[tab_name] is None:
                print(f"⚠️ Pestaña no disponible: {tab_name}")
                return

            if self.current_tab == tab_name:
                return

            # Ocultar pestaña actual
            if self.current_tab and self.current_tab in self.tabs and self.tabs[self.current_tab]:
                if hasattr(self.tabs[self.current_tab], 'hide'):
                    self.tabs[self.current_tab].hide()

            # Mostrar nueva pestaña
            if hasattr(self.tabs[tab_name], 'show'):
                self.tabs[tab_name].show()
                self.current_tab = tab_name

        except Exception as e:
            print(f"⚠️ Error mostrando pestaña {tab_name}: {e}")

    def on_closing(self):
        """Maneja el cierre de la aplicación simplificado."""
        try:
            print("🔄 Cerrando ContaFlow v2.0...")

            # Detener bot si está ejecutándose
            automatizacion_tab = self.tabs.get('automatizacion')
            if automatizacion_tab and hasattr(automatizacion_tab, 'bot_running'):
                if automatizacion_tab.bot_running:
                    print("⏹️ Deteniendo bot...")
                    automatizacion_tab.stop_bot()

                    # Esperar un momento para que se detenga correctamente
                    time.sleep(1)

            print("✅ ContaFlow v2.0 cerrado correctamente")
            self.destroy()

        except Exception as e:
            print(f"⚠️ Error durante cierre: {e}")
        finally:
            sys.exit(0)

    # ========== MÉTODOS DE INFORMACIÓN DEL SISTEMA ==========

    def get_system_info(self):
        """
        Obtiene información detallada del sistema simplificado.

        Returns:
            dict: Información del sistema
        """
        try:
            automatizacion_tab = self.tabs.get('automatizacion')

            system_info = {
                'version': '2.0',
                'system_name': 'ContaFlow - Sistema Simplificado',
                'system_type': 'simplified_cargador_search',
                'ui_framework': 'tkinter_native',
                'monitoring_method': 'Correos "Cargador" con archivos Excel',
                'monitoring_interval': '1 minuto (fijo)',
                'cache_system': 'Anti-duplicados habilitado',
                'search_robustness': 'Sin dependencia de estado UNSEEN',
                'tabs_available': list(self.tabs.keys()),
                'current_tab': self.current_tab,
                'features_removed': [
                    'Auto-inicio del bot',
                    'Configuración de intervalo variable',
                    'Búsqueda compleja por criterios',
                    'UI compleja con muchas opciones'
                ],
                'features_simplified': [
                    'Control básico del bot (start/stop)',
                    'Configuración mínima requerida',
                    'Búsqueda específica correos "Cargador"',
                    'Intervalo fijo optimizado'
                ]
            }

            # Agregar información del bot si está disponible
            if automatizacion_tab:
                bot_status = automatizacion_tab.get_bot_status()
                system_info.update({
                    'bot_available': True,
                    'bot_running': bot_status.get('running', False),
                    'bot_status': bot_status
                })
            else:
                system_info.update({
                    'bot_available': False,
                    'bot_running': False,
                    'bot_status': {}
                })

            return system_info

        except Exception as e:
            return {
                'version': '2.0',
                'system_name': 'ContaFlow - Sistema Simplificado',
                'error': str(e),
                'tabs_available': [],
                'current_tab': None,
                'bot_available': False
            }

    def get_configuration_status(self):
        """
        Obtiene el estado de configuración del sistema.

        Returns:
            dict: Estado de la configuración
        """
        try:
            # Intentar obtener información de configuración
            try:
                from contaflow.config.config_manager import ConfigManager
                config_manager = ConfigManager()
                config = config_manager.load_config()

                if not config:
                    return {
                        'configured': False,
                        'message': 'No hay configuración guardada',
                        'required_steps': [
                            'Configurar credenciales de email',
                            'Configurar carpeta de descarga'
                        ]
                    }

                # Verificar configuración básica
                email_configured = bool(config.get('email') and config.get('password'))
                search_configured = bool(config.get('search_criteria', {}).get('download_folder'))

                status = {
                    'configured': email_configured and search_configured,
                    'email_configured': email_configured,
                    'search_configured': search_configured,
                    'xml_configured': bool(config.get('xml_config')),
                    'recipients_configured': bool(config.get('recipients_config')),
                    'system_version': config.get('version', '2.0'),
                    'last_updated': config.get('last_updated', 'Desconocido')
                }

                # Generar recomendaciones
                recommendations = []
                if not email_configured:
                    recommendations.append('Configurar credenciales de email en pestaña Configuración')
                if not search_configured:
                    recommendations.append('Configurar carpeta de descarga en pestaña Configuración > Búsqueda')
                if not status['xml_configured']:
                    recommendations.append('Configurar procesamiento XML para funcionalidad completa (opcional)')
                if not status['recipients_configured']:
                    recommendations.append('Configurar destinatarios para envío automático (opcional)')

                status['recommendations'] = recommendations
                status['ready_to_start'] = email_configured and search_configured

                return status

            except ImportError:
                return {
                    'configured': False,
                    'error': 'ConfigManager no disponible',
                    'message': 'Error del sistema'
                }

        except Exception as e:
            return {
                'configured': False,
                'error': str(e),
                'message': 'Error verificando configuración'
            }

    def diagnose_system(self):
        """
        Realiza un diagnóstico completo del sistema.

        Returns:
            dict: Resultado del diagnóstico
        """
        diagnosis = {
            'timestamp': self._get_current_timestamp(),
            'system_health': 'unknown',
            'system_info': {},
            'configuration_status': {},
            'bot_status': {},
            'recommendations': [],
            'critical_issues': [],
            'warnings': []
        }

        try:
            # Obtener información del sistema
            diagnosis['system_info'] = self.get_system_info()

            # Obtener estado de configuración
            diagnosis['configuration_status'] = self.get_configuration_status()

            # Obtener estado del bot
            automatizacion_tab = self.tabs.get('automatizacion')
            if automatizacion_tab:
                diagnosis['bot_status'] = automatizacion_tab.get_bot_status()

            # Evaluar salud del sistema
            critical_issues = []
            warnings = []

            # Verificar pestañas críticas
            if not self.tabs.get('automatizacion'):
                critical_issues.append('Pestaña de Automatización no disponible')
            if not self.tabs.get('configuracion'):
                critical_issues.append('Pestaña de Configuración no disponible')

            # Verificar configuración
            config_status = diagnosis['configuration_status']
            if config_status.get('error'):
                critical_issues.append(f"Error en configuración: {config_status['error']}")
            elif not config_status.get('configured'):
                if not config_status.get('email_configured'):
                    critical_issues.append('Email no configurado')
                if not config_status.get('search_configured'):
                    critical_issues.append('Carpeta de descarga no configurada')

            # Verificar funcionalidades opcionales
            if not config_status.get('xml_configured'):
                warnings.append('Procesamiento XML no configurado - funcionalidad limitada')
            if not config_status.get('recipients_configured'):
                warnings.append('Envío automático no configurado')

            # Determinar salud general
            if critical_issues:
                diagnosis['system_health'] = 'critical' if len(critical_issues) > 2 else 'warning'
            elif warnings:
                diagnosis['system_health'] = 'healthy_with_warnings'
            else:
                diagnosis['system_health'] = 'healthy'

            diagnosis['critical_issues'] = critical_issues
            diagnosis['warnings'] = warnings

            # Generar recomendaciones
            recommendations = []
            if critical_issues:
                recommendations.extend([
                    'Resolver problemas críticos antes de usar el sistema',
                    'Verificar instalación de archivos del sistema'
                ])

            if config_status.get('recommendations'):
                recommendations.extend(config_status['recommendations'])

            if diagnosis['system_health'] == 'healthy':
                recommendations.append('Sistema listo para usar - puede iniciar el bot')

            diagnosis['recommendations'] = recommendations

            return diagnosis

        except Exception as e:
            diagnosis.update({
                'system_health': 'error',
                'error': str(e),
                'critical_issues': [f'Error en diagnóstico: {str(e)}'],
                'recommendations': ['Reiniciar aplicación', 'Verificar integridad de archivos']
            })
            return diagnosis

    def _get_current_timestamp(self):
        """Obtiene timestamp actual."""
        try:
            from datetime import datetime
            return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        except:
            return "timestamp_error"