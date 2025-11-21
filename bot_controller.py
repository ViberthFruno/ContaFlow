# bot_controller.py
"""
Controlador simplificado para la lógica de automatización de ContaFlow.
Maneja solo control básico del bot sin auto-inicio ni configuración de intervalo.
"""
# Archivos relacionados: automatizacion_ui.py, automatizacion_tab.py, config_manager.py

import threading
import time
from config_manager import ConfigManager


class BotController:
    """Clase simplificada que maneja el control básico del bot."""

    def __init__(self, parent, ui_interface):
        """
        Inicializa el controlador simplificado del bot.

        Args:
            parent: Widget padre (para after() calls)
            ui_interface: Instancia de AutomatizacionUI para comunicación
        """
        self.parent = parent
        self.ui = ui_interface
        self.config_manager = ConfigManager()

        # Estado del bot con protección thread-safe
        self._bot_lock = threading.Lock()
        self.bot_running = False
        self.bot_thread = None
        self.email_processor = None
        self.stop_event = threading.Event()

        # Flags para controlar estados críticos
        self.stopping_bot = False
        self.starting_bot = False
        self._last_action_time = 0
        self._action_debounce = 1.0  # 1 segundo entre acciones críticas

    def initialize(self):
        """Inicializa el controlador simplificado."""
        try:
            # Configurar callbacks simplificados en la UI
            self.ui.set_callbacks(
                toggle_bot_callback=self.handle_toggle_bot_click,
                clear_log_callback=self.handle_clear_log_click
            )

            self._log_debug("BotController simplificado inicializado correctamente")

        except Exception as e:
            self._log_debug(f"Error inicializando BotController: {e}")

    # ========== MÉTODOS DE CONTROL DEL BOT ==========

    def handle_toggle_bot_click(self):
        """
        Maneja el clic en el botón de toggle con debounce y protección.
        """
        current_time = time.time()

        # Implementar debounce para acciones críticas
        if current_time - self._last_action_time < self._action_debounce:
            self._log_debug("Debounce activo, ignorando clic de toggle")
            return

        self._last_action_time = current_time

        # Llamar al método de toggle
        self.toggle_bot()

    def toggle_bot(self):
        """Inicia o detiene el bot con manejo robusto."""
        with self._bot_lock:
            # Verificar si ya estamos en una transición
            if self.stopping_bot:
                self.ui.add_log_message("⚠️ El bot ya se está deteniendo, espere...", "warning")
                return

            if self.starting_bot:
                self.ui.add_log_message("⚠️ El bot ya se está iniciando, espere...", "warning")
                return

            # Determinar acción basada en el estado actual
            if not self.bot_running:
                self.start_bot()
            else:
                self.stop_bot()

    def start_bot(self):
        """Inicia el bot de automatización con configuración simplificada."""
        if self.starting_bot or self.bot_running:
            self._log_debug("Bot ya está iniciando o corriendo")
            return

        self.starting_bot = True
        self._log_debug("Iniciando proceso de arranque del bot")

        try:
            # Verificar configuración antes de iniciar
            if not self.verify_configuration():
                self.ui.add_log_message("❌ No se puede iniciar: Configuración incompleta", "error")
                return

            # Importar aquí para evitar dependencias circulares
            from email_processor import EmailProcessor

            # Limpiar evento de parada anterior y resetear estados
            self.stop_event.clear()
            self.stopping_bot = False

            self._log_debug("Configuración verificada, creando EmailProcessor")

            # Crear instancia del procesador
            self.email_processor = EmailProcessor(self)

            # Iniciar thread del bot
            self.bot_thread = threading.Thread(
                target=self._bot_thread_wrapper,
                daemon=True
            )
            self.bot_thread.start()

            self._log_debug("Thread del bot iniciado")

            # Actualizar estado
            self.bot_running = True
            self.ui.update_ui_for_running_state()

            self.ui.add_log_message("✅ Bot iniciado - Monitoreo cada 1 minuto", "success")
            self.ui.add_log_message("🎯 Buscando correos 'Cargador' con archivos Excel", "info")

        except ImportError as e:
            self.ui.add_log_message("❌ Error: No se pudo cargar el procesador de emails", "error")
            self._log_debug(f"ImportError: {e}")
        except Exception as e:
            self.ui.add_log_message(f"❌ Error al iniciar bot: {str(e)}", "error")
            self._log_debug(f"Exception en start_bot: {e}")
            # Limpiar estado en caso de error
            self._cleanup_failed_start()

        finally:
            self.starting_bot = False

    def stop_bot(self):
        """Detiene el bot de automatización de forma robusta y no-bloqueante."""
        if self.stopping_bot or not self.bot_running:
            self._log_debug("Bot ya está deteniéndose o no está corriendo")
            return

        try:
            self.stopping_bot = True
            self._log_debug("Iniciando proceso de parada del bot")

            self.ui.add_log_message("⏹️ Deteniendo bot...", "info")

            # Actualizar UI inmediatamente para mostrar que se está deteniendo
            self.ui.update_ui_for_stopping_state()

            # Señalar parada inmediatamente
            self.stop_event.set()

            # Detener el procesador si existe
            if self.email_processor:
                try:
                    self.email_processor.stop_monitoring()
                    self._log_debug("stop_monitoring() llamado en EmailProcessor")
                except Exception as e:
                    self._log_debug(f"Error llamando stop_monitoring(): {e}")

            # Usar un thread separado para la limpieza sin bloquear la GUI
            cleanup_thread = threading.Thread(target=self._cleanup_bot_thread_safe, daemon=True)
            cleanup_thread.start()

        except Exception as e:
            self.ui.add_log_message(f"❌ Error al detener bot: {str(e)}", "error")
            self._log_debug(f"Excepción en stop_bot: {e}")
            self._reset_stop_state()

    def force_stop_bot(self):
        """
        Fuerza la parada del bot en caso de problemas.
        Método de emergencia.
        """
        self._log_debug("Forzando parada del bot")

        with self._bot_lock:
            self.stop_event.set()
            self.stopping_bot = True

            # Limpiar referencias inmediatamente
            if self.email_processor:
                try:
                    self.email_processor.stop_monitoring()
                except:
                    pass

            self.email_processor = None

            # Resetear estados
            self.bot_running = False
            self.stopping_bot = False
            self.starting_bot = False

            # Actualizar UI
            self.ui.update_ui_for_stopped_state()
            self.ui.add_log_message("🔧 Bot forzado a detenerse", "warning")

    # ========== MÉTODOS DE THREAD MANAGEMENT ==========

    def _bot_thread_wrapper(self):
        """
        Wrapper para el thread del bot con manejo de excepciones.
        """
        try:
            self._log_debug("Ejecutando email_processor.start_monitoring()")
            self.email_processor.start_monitoring()
        except Exception as e:
            self._log_debug(f"Excepción en bot thread: {e}")
            # Notificar a la UI del error desde el hilo principal
            self.parent.after(0, lambda: self.ui.add_log_message(f"❌ Error en hilo del bot: {str(e)}", "error"))
            self.parent.after(0, self._handle_bot_thread_error)

    def _handle_bot_thread_error(self):
        """Maneja errores del thread del bot en el hilo principal."""
        with self._bot_lock:
            if self.bot_running:
                self._log_debug("Manejando error del thread del bot")
                self.bot_running = False
                self.stopping_bot = False
                self.ui.update_ui_for_stopped_state()
                self.ui.add_log_message("❌ El bot se detuvo debido a un error", "error")

    def _cleanup_bot_thread_safe(self):
        """Limpia el thread del bot sin bloquear la GUI de forma segura."""
        try:
            self._log_debug("Iniciando limpieza del thread del bot")

            # Esperar que el thread termine con timeout
            if self.bot_thread and self.bot_thread.is_alive():
                self._log_debug("Esperando que termine el thread del bot")
                self.bot_thread.join(timeout=3.0)

                # Si el thread sigue vivo después del timeout, loggear warning
                if self.bot_thread.is_alive():
                    self._log_debug("Thread del bot no terminó en el timeout, forzando cierre")
                    self.parent.after(0, lambda: self.ui.add_log_message("⚠️ Forzando cierre del bot...", "warning"))

            # Actualizar UI desde el hilo principal
            self.parent.after(0, self._finish_bot_stop_safe)

        except Exception as e:
            self._log_debug(f"Excepción en _cleanup_bot_thread_safe: {e}")
            # Actualizar UI con error desde el hilo principal
            self.parent.after(0, lambda: self.ui.add_log_message(f"❌ Error en limpieza: {str(e)}", "error"))
            self.parent.after(0, self._reset_stop_state)

    def _finish_bot_stop_safe(self):
        """Finaliza el proceso de parada del bot de forma segura (ejecutado en el hilo principal)."""
        with self._bot_lock:
            try:
                self._log_debug("Finalizando parada del bot")

                # Limpiar referencias
                if self.email_processor:
                    try:
                        self.email_processor = None
                    except:
                        pass

                self.bot_thread = None

                # Actualizar estado
                self.bot_running = False
                self.stopping_bot = False

                # Actualizar UI
                self.ui.update_ui_for_stopped_state()

                self.ui.add_log_message("✅ Bot detenido correctamente", "success")
                self._log_debug("Bot detenido correctamente")

            except Exception as e:
                self._log_debug(f"Excepción en _finish_bot_stop_safe: {e}")
                self.ui.add_log_message(f"❌ Error finalizando parada: {str(e)}", "error")
                self._reset_stop_state()

    def _cleanup_failed_start(self):
        """Limpia el estado después de un arranque fallido."""
        self.bot_running = False
        self.stopping_bot = False
        self.starting_bot = False
        if self.email_processor:
            try:
                self.email_processor.stop_monitoring()
            except:
                pass
        self.email_processor = None
        self.bot_thread = None

    def _reset_stop_state(self):
        """Resetea el estado de parada en caso de error."""
        with self._bot_lock:
            self.stopping_bot = False
            self.starting_bot = False
            self.ui.reset_stop_state(self.bot_running)

    # ========== MÉTODOS DE CALLBACKS SIMPLIFICADOS ==========

    def handle_clear_log_click(self):
        """Maneja el clic en el botón de limpiar log."""
        self.ui.clear_log()

    # ========== MÉTODOS DE CONFIGURACIÓN ==========

    def verify_configuration(self):
        """
        Verifica que la configuración sea completa para el sistema simplificado.

        Returns:
            bool: True si la configuración es válida
        """
        try:
            config = self.config_manager.load_config()

            if not config:
                self._log_debug("No hay configuración cargada")
                return False

            # Verificar configuración de email
            email_fields = ['provider', 'email', 'password']
            for field in email_fields:
                if not config.get(field):
                    self._log_debug(f"Campo de email faltante: {field}")
                    return False

            # Verificar configuración de búsqueda (solo carpeta de descarga)
            search_criteria = config.get('search_criteria')
            if not search_criteria:
                self._log_debug("Criterios de búsqueda faltantes")
                return False

            if not search_criteria.get('download_folder'):
                self._log_debug("Carpeta de descarga faltante")
                return False

            self._log_debug("Configuración verificada correctamente")
            return True

        except Exception as e:
            self._log_debug(f"Error verificando configuración: {e}")
            return False

    # ========== MÉTODOS DE ESTADO ==========

    def get_bot_status(self):
        """
        Obtiene el estado actual del bot de forma thread-safe.

        Returns:
            dict: Estado del bot
        """
        with self._bot_lock:
            return {
                'running': self.bot_running,
                'stopping': self.stopping_bot,
                'starting': self.starting_bot,
                'has_thread': self.bot_thread is not None,
                'thread_alive': self.bot_thread.is_alive() if self.bot_thread else False,
                'has_processor': self.email_processor is not None,
                'system_type': 'simplified_cargador_search',
                'monitoring_interval': '1_minute_fixed'
            }

    # ========== MÉTODOS DE COMPATIBILIDAD CON EMAIL_PROCESSOR ==========

    def add_log_message(self, message, msg_type="info"):
        """
        Proxy method para compatibilidad con email_processor.
        Redirige a la UI.

        Args:
            message (str): Mensaje a agregar
            msg_type (str): Tipo de mensaje
        """
        self.ui.add_log_message(message, msg_type)

    def update_statistics(self, emails_found=None, files_downloaded=None):
        """
        Método mantenido para compatibilidad con email_processor.
        Redirige a la UI.

        Args:
            emails_found (int): Número de emails encontrados
            files_downloaded (int): Número de archivos descargados
        """
        self.ui.update_statistics(emails_found, files_downloaded)

    # ========== MÉTODOS DE LOGGING INTERNO ==========

    def _log_debug(self, message):
        """
        Método de logging discreto para debugging interno.

        Args:
            message (str): Mensaje a loggear
        """
        # Para debugging - puede ser comentado en producción
        # print(f"[DEBUG] BotController: {message}")
        pass

    # ========== MÉTODO DE LIMPIEZA ==========

    def cleanup(self):
        """Limpia recursos del controlador de forma segura."""
        try:
            if self.bot_running:
                self._log_debug("Cleanup: deteniendo bot")
                self.force_stop_bot()
        except Exception as e:
            self._log_debug(f"Error en cleanup: {e}")
        finally:
            # Asegurar limpieza
            try:
                if hasattr(self, 'stop_event'):
                    self.stop_event.set()
            except:
                pass

    # ========== MÉTODOS DE INFORMACIÓN DEL SISTEMA ==========

    def get_system_capabilities(self):
        """
        Obtiene las capacidades del sistema simplificado.

        Returns:
            dict: Capacidades disponibles
        """
        return {
            'search_method': 'Correos con asunto "Cargador"',
            'file_types': 'Solo archivos Excel (.xlsx, .xls)',
            'monitoring_interval': '1 minuto (fijo - no configurable)',
            'auto_start': 'No disponible (eliminado)',
            'cache_system': 'Sistema anti-duplicados habilitado',
            'thread_management': 'Seguro y robusto',
            'configuration_required': [
                'Credenciales de email',
                'Carpeta de descarga'
            ],
            'optional_features': [
                'Procesamiento XML',
                'Envío automático de resultados'
            ],
            'system_improvements': [
                'Búsqueda robusta sin estado UNSEEN',
                'Intervalo de monitoreo fijo optimizado',
                'UI simplificada sin opciones innecesarias',
                'Eliminación completa de auto-inicio',
                'Control de bot simplificado y confiable'
            ]
        }