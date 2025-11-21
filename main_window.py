# main_window.py
"""
Ventana principal simplificada de ContaFlow usando tkinter nativo.
Interfaz limpia con navegación por pestañas sin auto-inicio del bot.
"""
# Archivos relacionados: automatizacion_tab.py, configuracion_tab.py

import tkinter as tk
from tkinter import ttk, messagebox
import tkinter.font as tkfont
import sys
import threading
import time


class MainWindow(tk.Tk):
    """Ventana principal simplificada con tkinter nativo sin auto-inicio."""

    def __init__(self):
        """Inicializa la ventana principal simplificada."""
        super().__init__()
        print("🏗️ Inicializando ventana principal de ContaFlow v2.0...")

        # Variables de control simplificadas
        self.tabs = {}
        self.current_tab = None

        # Configurar fuentes globales
        self.setup_fonts()

        # Configurar la ventana
        self.setup_window()

        # Crear interfaz
        self.create_interface()

        # Inicializar pestañas
        self.initialize_tabs()

        # Mostrar pestaña por defecto
        self.show_tab("automatizacion")

        print("✅ ContaFlow v2.0 - Sistema Simplificado iniciado correctamente")

    def setup_fonts(self):
        """Configura las fuentes globales de la aplicación."""
        try:
            # Fuente por defecto
            default_font = tkfont.nametofont("TkDefaultFont")
            default_font.configure(family="Arial", size=10)

            # Fuente de texto
            text_font = tkfont.nametofont("TkTextFont")
            text_font.configure(family="Arial", size=10)

            # Fuente de menú
            menu_font = tkfont.nametofont("TkMenuFont")
            menu_font.configure(family="Arial", size=10)

            print("✓ Fuentes configuradas correctamente")
        except Exception as e:
            print(f"⚠️ Error configurando fuentes: {e}")

    def setup_window(self):
        """Configura las propiedades básicas de la ventana."""
        try:
            # Título y dimensiones
            self.title("API iFR Pro + Bot de Correo - Sistema Integrado")
            self.geometry("900x600")
            self.configure(bg="#f0f0f0")  # Fondo gris claro
            self.minsize(800, 500)

            # Centrar ventana
            self._center_window()

            # Configurar cierre
            self.protocol("WM_DELETE_WINDOW", self.on_closing)

            # Configurar icono (opcional)
            try:
                self.iconbitmap("icon.ico")
            except:
                pass

        except Exception as e:
            print(f"⚠️ Error configurando ventana: {e}")

    def _center_window(self):
        """Centra la ventana en la pantalla usando el método estándar."""
        try:
            self.update_idletasks()
            width = self.winfo_width()
            height = self.winfo_height()
            x = (self.winfo_screenwidth() // 2) - (width // 2)
            y = (self.winfo_screenheight() // 2) - (height // 2)
            self.geometry(f'{width}x{height}+{x}+{y}')
        except Exception as e:
            print(f"⚠️ Error centrando ventana: {e}")

    def center_window(self):
        """Alias para compatibilidad con código existente."""
        self._center_window()

    def create_interface(self):
        """Crea la interfaz principal con notebook de pestañas."""
        # Frame principal con configuración de grid
        main_frame = ttk.Frame(self, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)

        # Título principal grande con mejor estilo
        title_frame = ttk.Frame(main_frame)
        title_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        title_label = ttk.Label(
            title_frame,
            text="🤖 Bot ContaFlow - Sistema Integrado",
            font=("Arial", 18, "bold"),
            foreground="#2c3e50"
        )
        title_label.pack()

        subtitle_label = ttk.Label(
            title_frame,
            text="API iFR Pro + Automatización de Correos",
            font=("Arial", 10),
            foreground="#7f8c8d"
        )
        subtitle_label.pack()

        # Notebook para pestañas con mejor configuración
        self.notebook = ttk.Notebook(main_frame)
        self.notebook.grid(row=1, column=0, sticky="nsew")

        # Crear frames para cada pestaña
        self.automatizacion_frame = ttk.Frame(self.notebook, padding="10")
        self.configuracion_frame = ttk.Frame(self.notebook, padding="10")

        # Agregar pestañas al notebook con mejor texto
        self.notebook.add(self.automatizacion_frame, text="⚡ Panel Principal")
        self.notebook.add(self.configuracion_frame, text="⚙️ Configuración")

        # Vincular evento de cambio de pestaña
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

    def initialize_tabs(self):
        """Inicializa las pestañas del sistema simplificado."""
        try:
            # Importar y crear pestaña de automatización
            from automatizacion_tab import AutomatizacionTab
            self.tabs["automatizacion"] = AutomatizacionTab(self.automatizacion_frame)
            print("✅ Pestaña de automatización inicializada - Sistema simplificado")
        except Exception as e:
            print(f"⚠️ Error inicializando automatización: {e}")
            self.tabs["automatizacion"] = None

        try:
            # Importar y crear pestaña de configuración
            from configuracion_tab import ConfiguracionTab
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
        """Maneja el cierre de la aplicación con confirmación modal."""
        try:
            # Verificar si el bot está corriendo
            automatizacion_tab = self.tabs.get('automatizacion')
            bot_running = False

            if automatizacion_tab and hasattr(automatizacion_tab, 'bot_running'):
                bot_running = automatizacion_tab.bot_running

            # Si el bot está corriendo, pedir confirmación
            if bot_running:
                if not messagebox.askyesno(
                    "Confirmar Cierre",
                    "El bot está en ejecución.\n\n¿Está seguro que desea cerrar la aplicación?",
                    icon='warning'
                ):
                    return

            print("🔄 Cerrando ContaFlow v2.0...")

            # Detener bot si está ejecutándose
            if bot_running and automatizacion_tab:
                print("⏹️ Deteniendo bot...")
                automatizacion_tab.stop_bot()

                # Esperar un momento para que se detenga correctamente
                time.sleep(1)

            print("✅ ContaFlow v2.0 cerrado correctamente")
            self.quit()
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
                from config_manager import ConfigManager
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