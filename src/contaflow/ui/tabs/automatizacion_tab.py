# automatizacion_tab.py
"""
Coordinador simplificado de automatización para ContaFlow.
Integra AutomatizacionUI (tkinter nativo) y BotController sin auto-inicio ni configuración de intervalo.
"""
# Archivos relacionados: automatizacion_ui.py, bot_controller.py, config_manager.py

from contaflow.ui.tabs.automatizacion_ui import AutomatizacionUI
from contaflow.core.bot_controller import BotController
from contaflow.config.config_manager import ConfigManager


class AutomatizacionTab:
    """Coordinador simplificado que integra UI (tkinter nativo) y Controller."""

    def __init__(self, parent):
        """
        Inicializa la pestaña de automatización simplificada.

        Args:
            parent: Widget padre donde se colocará esta pestaña
        """
        self.parent = parent
        self.config_manager = ConfigManager()

        # Crear la interfaz de usuario simplificada
        self.ui = AutomatizacionUI(parent, self.config_manager)

        # Crear el controlador de lógica simplificado
        self.controller = BotController(parent, self.ui)

        # Inicializar componentes
        self._initialize_components()

    def _initialize_components(self):
        """Inicializa y conecta los componentes simplificados."""
        try:
            # Crear la interfaz
            self.ui.create_interface()

            # Inicializar el controlador (configura callbacks simplificados)
            self.controller.initialize()

            self._log_debug("AutomatizacionTab simplificada inicializada correctamente")

        except Exception as e:
            self._log_debug(f"Error inicializando AutomatizacionTab: {e}")

    # ========== PROPIEDADES SIMPLIFICADAS ==========

    @property
    def is_visible(self):
        """Obtiene el estado de visibilidad."""
        return self.ui.is_visible

    @property
    def bot_running(self):
        """Obtiene el estado del bot."""
        return self.controller.bot_running

    @property
    def stopping_bot(self):
        """Obtiene si el bot se está deteniendo."""
        return self.controller.stopping_bot

    @property
    def starting_bot(self):
        """Obtiene si el bot se está iniciando."""
        return self.controller.starting_bot

    # ========== MÉTODOS DE VISIBILIDAD ==========

    def show(self):
        """Muestra la pestaña de automatización."""
        try:
            self.ui.show()
            self._log_debug("Pestaña de automatización mostrada")
        except Exception as e:
            self._log_debug(f"Error mostrando pestaña de automatización: {e}")

    def hide(self):
        """Oculta la pestaña de automatización."""
        try:
            self.ui.hide()
            self._log_debug("Pestaña de automatización ocultada")
        except Exception as e:
            self._log_debug(f"Error ocultando pestaña de automatización: {e}")

    # ========== MÉTODOS DE CONTROL DEL BOT ==========

    def start_bot(self):
        """Inicia el bot de automatización."""
        self.controller.start_bot()

    def stop_bot(self):
        """Detiene el bot de automatización."""
        self.controller.stop_bot()

    def toggle_bot(self):
        """Alterna el estado del bot (iniciar/detener)."""
        self.controller.toggle_bot()

    def force_stop_bot(self):
        """Fuerza la parada del bot en caso de emergencia."""
        self.controller.force_stop_bot()

    # ========== MÉTODOS DE LOGGING ==========

    def add_log_message(self, message, msg_type="info"):
        """
        Agrega un mensaje al log.

        Args:
            message (str): Mensaje a agregar
            msg_type (str): Tipo de mensaje (info, success, error, warning)
        """
        self.ui.add_log_message(message, msg_type)

    def clear_log(self):
        """Limpia el log de actividad."""
        self.ui.clear_log()

    # ========== MÉTODOS DE CONFIGURACIÓN ==========

    def verify_configuration(self):
        """
        Verifica que la configuración sea completa.

        Returns:
            bool: True si la configuración es válida
        """
        return self.controller.verify_configuration()

    # ========== MÉTODOS DE ESTADO ==========

    def get_bot_status(self):
        """
        Obtiene el estado actual del bot de forma thread-safe.

        Returns:
            dict: Estado del bot
        """
        return self.controller.get_bot_status()

    def update_bot_status(self, status_text, color):
        """
        Actualiza el estado visual del bot.

        Args:
            status_text (str): Texto del estado
            color (str): Color del texto (tkinter nativo)
        """
        # Convertir colores CustomTkinter a tkinter nativo si es necesario
        color_map = {
            "#4CAF50": "green",
            "#f44336": "red",
            "#FFC107": "orange",
            "#2196F3": "blue"
        }
        native_color = color_map.get(color, color)
        self.ui.update_bot_status(status_text, native_color)

    # ========== MÉTODOS DE COMPATIBILIDAD ==========

    def update_statistics(self, emails_found=None, files_downloaded=None):
        """Método mantenido para compatibilidad con email_processor."""
        self.ui.update_statistics(emails_found, files_downloaded)

    # ========== MÉTODOS INTERNOS ==========

    def _log_debug(self, message):
        """
        Método de logging discreto para debugging interno.

        Args:
            message (str): Mensaje a loggear
        """
        # Para debugging - puede ser comentado en producción
        # print(f"[DEBUG] AutomatizacionTab: {message}")
        pass

    # ========== MÉTODO DE LIMPIEZA ==========

    def __del__(self):
        """Destructor para limpiar recursos de forma segura."""
        try:
            if hasattr(self, 'controller'):
                self._log_debug("Destructor: limpiando controlador")
                self.controller.cleanup()
        except Exception as e:
            self._log_debug(f"Error en destructor: {e}")

    # ========== MÉTODOS DE ACCESO DIRECTO A COMPONENTES ==========

    def get_ui_component(self):
        """
        Obtiene referencia al componente UI.

        Returns:
            AutomatizacionUI: Instancia del componente UI simplificado
        """
        return self.ui

    def get_controller_component(self):
        """
        Obtiene referencia al componente Controller.

        Returns:
            BotController: Instancia del componente Controller simplificado
        """
        return self.controller

    # ========== MÉTODOS DE INFORMACIÓN ==========

    def get_component_info(self):
        """
        Obtiene información sobre los componentes internos.

        Returns:
            dict: Información de los componentes
        """
        try:
            return {
                'ui_created': hasattr(self, 'ui') and self.ui is not None,
                'controller_created': hasattr(self, 'controller') and self.controller is not None,
                'ui_visible': self.ui.is_visible if hasattr(self, 'ui') else False,
                'bot_status': self.controller.get_bot_status() if hasattr(self, 'controller') else {},
                'config_manager_available': hasattr(self, 'config_manager') and self.config_manager is not None,
                'ui_framework': 'tkinter_native',
                'system_type': 'simplified_cargador_search',
                'monitoring_interval': '1_minute_fixed',
                'auto_start_available': False,  # Eliminado completamente
                'features_removed': ['auto_start', 'configurable_interval']
            }
        except Exception as e:
            return {
                'error': str(e),
                'ui_created': False,
                'controller_created': False,
                'ui_visible': False,
                'bot_status': {},
                'config_manager_available': False,
                'ui_framework': 'unknown',
                'system_type': 'error'
            }

    def get_system_info(self):
        """
        Obtiene información del sistema simplificado.

        Returns:
            dict: Información del sistema
        """
        try:
            config = self.config_manager.load_config()

            return {
                'system_version': 'ContaFlow v2.0 - Sistema Simplificado',
                'search_target': 'Correos con asunto "Cargador"',
                'file_types': 'Solo archivos Excel (.xlsx, .xls)',
                'monitoring_interval': '1 minuto (fijo)',
                'search_method': 'Búsqueda robusta sin estado UNSEEN',
                'auto_start': 'Eliminado (no disponible)',
                'cache_system': 'Sistema anti-duplicados habilitado',
                'configuration_status': {
                    'email_configured': bool(config and config.get('email')),
                    'search_configured': bool(config and config.get('search_criteria', {}).get('download_folder')),
                    'xml_configured': bool(config and config.get('xml_config')),
                    'recipients_configured': bool(config and config.get('recipients_config'))
                },
                'features_simplified': [
                    'Eliminado auto-inicio',
                    'Intervalo fijo 1 minuto',
                    'Búsqueda solo "Cargador"',
                    'UI simplificada',
                    'Configuración mínima'
                ]
            }
        except Exception as e:
            return {
                'system_version': 'ContaFlow v2.0 - Error',
                'error': str(e)
            }

    # ========== MÉTODOS DE DIAGNÓSTICO ==========

    def diagnose_system(self):
        """
        Realiza diagnóstico del sistema simplificado.

        Returns:
            dict: Resultado del diagnóstico
        """
        try:
            diagnosis = {
                'timestamp': self._get_current_timestamp(),
                'system_health': 'checking',
                'components': {},
                'configuration': {},
                'recommendations': []
            }

            # Verificar componentes
            diagnosis['components'] = {
                'ui_available': hasattr(self, 'ui') and self.ui is not None,
                'controller_available': hasattr(self, 'controller') and self.controller is not None,
                'config_manager_available': hasattr(self, 'config_manager') and self.config_manager is not None,
                'bot_running': self.bot_running if hasattr(self, 'controller') else False
            }

            # Verificar configuración
            try:
                config = self.config_manager.load_config()
                diagnosis['configuration'] = {
                    'config_exists': config is not None,
                    'email_configured': bool(config and config.get('email')),
                    'search_folder_configured': bool(
                        config and config.get('search_criteria', {}).get('download_folder')),
                    'xml_processing_available': bool(config and config.get('xml_config')),
                    'email_sending_available': bool(config and config.get('recipients_config'))
                }
            except Exception as e:
                diagnosis['configuration'] = {'error': str(e)}

            # Generar recomendaciones
            if not diagnosis['configuration'].get('email_configured'):
                diagnosis['recommendations'].append("Configurar credenciales de email")

            if not diagnosis['configuration'].get('search_folder_configured'):
                diagnosis['recommendations'].append("Configurar carpeta de descarga en pestaña Búsqueda")

            if not diagnosis['configuration'].get('xml_processing_available'):
                diagnosis['recommendations'].append("Configurar procesamiento XML para funcionalidad completa")

            # Determinar salud general
            critical_issues = []
            if not diagnosis['components']['ui_available']:
                critical_issues.append("UI no disponible")
            if not diagnosis['components']['controller_available']:
                critical_issues.append("Controller no disponible")
            if not diagnosis['configuration'].get('email_configured'):
                critical_issues.append("Email no configurado")
            if not diagnosis['configuration'].get('search_folder_configured'):
                critical_issues.append("Carpeta de descarga no configurada")

            if not critical_issues:
                diagnosis['system_health'] = 'healthy'
            elif len(critical_issues) <= 2:
                diagnosis['system_health'] = 'warning'
            else:
                diagnosis['system_health'] = 'critical'

            diagnosis['critical_issues'] = critical_issues

            return diagnosis

        except Exception as e:
            return {
                'timestamp': self._get_current_timestamp(),
                'system_health': 'error',
                'error': str(e),
                'components': {},
                'configuration': {},
                'recommendations': ['Reiniciar aplicación', 'Verificar archivos del sistema']
            }

    def _get_current_timestamp(self):
        """Obtiene timestamp actual."""
        try:
            from datetime import datetime
            return datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        except:
            return "timestamp_error"