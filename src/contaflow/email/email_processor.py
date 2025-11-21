# email_processor.py
"""
Procesador principal simplificado de emails para ContaFlow con monitoreo fijo de 1 minuto.
Búsqueda robusta de correos 'Cargador' con Excel, procesamiento automático y sistema robusto.
"""
# Archivos relacionados: email_manager.py, config_manager.py, excel_processor.py, email_sender.py

import time
import threading
from datetime import datetime
from contaflow.email.email_manager import EmailManager
from contaflow.config.config_manager import ConfigManager


class EmailProcessor:
    """Clase simplificada para procesamiento automático robusto de emails con intervalo fijo."""

    def __init__(self, automation_tab):
        self.automation_tab = automation_tab
        self.email_manager = EmailManager()
        self.config_manager = ConfigManager()

        # Control de procesamiento
        self.is_running = False
        self.monitoring_thread = None
        self.stop_event = threading.Event()

        # Configuración cargada
        self.email_config = None
        self.xml_config = None
        self.recipients_config = None
        self.download_folder = None

        # Componentes de procesamiento
        self.excel_processor = None
        self.email_sender = None

        # Estadísticas simplificadas
        self.session_stats = {
            'cycles_completed': 0,
            'emails_found': 0,
            'files_downloaded': 0,
            'files_processed': 0,
            'emails_sent': 0,
            'errors': 0,
            'connection_retries': 0,
            'last_check_time': None
        }

        # Configuración fija
        self.MONITORING_INTERVAL_MINUTES = 1  # FIJO: 1 minuto
        self.max_retries = 3
        self.retry_delay = 5
        self.stop_check_interval = 0.2

    def start_monitoring(self):
        """Inicia el monitoreo automático con intervalo fijo de 1 minuto."""
        try:
            self.is_running = True
            self.stop_event.clear()

            self.log_message("🚀 Iniciando monitoreo robusto de correos 'Cargador'", "info")
            self.log_message(f"⏰ Intervalo fijo: {self.MONITORING_INTERVAL_MINUTES} minuto", "info")

            if not self._load_configuration() or not self._initialize_processors() or not self._establish_connection():
                self.log_message("❌ Error: No se pudo inicializar el sistema", "error")
                return

            self.log_message("✅ Monitoreo robusto iniciado correctamente", "success")
            self._monitoring_loop()

        except Exception as e:
            self.log_message(f"❌ Error crítico en monitoreo: {str(e)}", "error")
            self.session_stats['errors'] += 1
        finally:
            self._cleanup()

    def stop_monitoring(self):
        """Detiene el monitoreo de correos."""
        self.is_running = False
        self.stop_event.set()
        self.log_message("⏹️ Señal de parada enviada al procesador", "info")

        # Detener componentes
        for processor in [self.excel_processor, self.email_sender]:
            if processor:
                try:
                    if hasattr(processor, 'stop_processing'):
                        processor.stop_processing()
                    elif hasattr(processor, 'stop_sending'):
                        processor.stop_sending()
                except:
                    pass

    def _load_configuration(self):
        """Carga configuración simplificada."""
        try:
            self.log_message("📋 Cargando configuración...", "info")
            config = self.config_manager.load_config()

            if not config:
                self.log_message("❌ No se encontró configuración", "error")
                return False

            # Validar configuración de email
            required_fields = ['provider', 'email', 'password']
            if not all(config.get(field) for field in required_fields):
                self.log_message("❌ Campos de email incompletos", "error")
                return False

            self.email_config = {field: config[field] for field in required_fields}

            # Configuración de descarga (simplificada)
            search_criteria = config.get('search_criteria')
            if not search_criteria or not search_criteria.get('download_folder'):
                self.log_message("❌ Carpeta de descarga no configurada", "error")
                return False

            self.download_folder = search_criteria['download_folder']

            # Cargar configuraciones opcionales
            self.xml_config = config.get('xml_config')
            self.recipients_config = config.get('recipients_config')

            # Log configuración cargada
            self.log_message(f"✅ Email configurado: {self.email_config['email']}", "success")
            self.log_message(f"✅ Carpeta de descarga: {self.download_folder}", "success")

            if self.xml_config:
                companies = self._get_configured_companies_count()
                self.log_message(f"✅ Procesamiento XML habilitado: {companies} empresas", "success")

            if self.recipients_config:
                cc_count = len(self.recipients_config.get('cc_recipients', []))
                self.log_message(f"✅ Envío automático habilitado: 1 principal + {cc_count} CC", "success")

            return True

        except Exception as e:
            self.log_message(f"❌ Error cargando configuración: {str(e)}", "error")
            return False

    def _get_configured_companies_count(self):
        """Obtiene el número de empresas configuradas."""
        if not self.xml_config:
            return 0

        if 'company_folders' in self.xml_config:
            return len(self.xml_config['company_folders'])
        elif 'xml_folder' in self.xml_config:
            return 1

        return 0

    def _initialize_processors(self):
        """Inicializa componentes de procesamiento."""
        try:
            # Inicializar ExcelProcessor si hay configuración XML
            if self.xml_config:
                try:
                    from contaflow.core.excel_processor import ExcelProcessor
                    self.excel_processor = ExcelProcessor(self.automation_tab)
                    self.log_message("✅ Procesador Excel/XML inicializado", "success")
                except (ImportError, Exception) as e:
                    self.log_message(f"⚠️ ExcelProcessor no disponible: {str(e)}", "warning")
                    self.xml_config = None

            # Inicializar EmailSender si hay configuración de destinatarios
            if self.recipients_config:
                try:
                    from contaflow.email.email_sender import EmailSender
                    self.email_sender = EmailSender(self.automation_tab)
                    self.log_message("✅ Sistema de envío inicializado", "success")
                except (ImportError, Exception) as e:
                    self.log_message(f"⚠️ EmailSender no disponible: {str(e)}", "warning")
                    self.recipients_config = None

            # Determinar modo de operación
            modes = []
            if self.xml_config:
                modes.append("Procesamiento XML")
            if self.recipients_config:
                modes.append("Envío automático")

            mode_text = " + ".join(["Descarga"] + modes) if modes else "Solo descarga"
            self.log_message(f"🔄 Modo de operación: {mode_text}", "info")

            return True

        except Exception as e:
            self.log_message(f"❌ Error inicializando procesadores: {str(e)}", "error")
            return False

    def _establish_connection(self):
        """Establece conexión con el servidor de correo."""
        for attempt in range(self.max_retries):
            if self.stop_event.is_set():
                return False

            try:
                self.log_message(f"🔌 Conectando al servidor (intento {attempt + 1}/{self.max_retries})...", "info")

                if self.email_manager.connect(self.email_config['provider'],
                                              self.email_config['email'],
                                              self.email_config['password']):
                    self.log_message(f"✅ Conectado a {self.email_config['provider']}", "success")
                    return True
                else:
                    self.log_message(f"❌ Error de autenticación (intento {attempt + 1})", "error")

            except Exception as e:
                self.log_message(f"❌ Error conectando (intento {attempt + 1}): {str(e)}", "error")

            self.session_stats['connection_retries'] += 1

            if attempt < self.max_retries - 1 and not self._wait_with_interrupt(self.retry_delay):
                continue

        self.log_message("❌ No se pudo establecer conexión", "error")
        return False

    def _monitoring_loop(self):
        """Ciclo principal simplificado con intervalo fijo."""
        consecutive_errors = 0
        max_consecutive_errors = 5

        while self.is_running and not self.stop_event.is_set():
            try:
                if self.stop_event.is_set():
                    break

                # Verificar conexión
                if not self._ensure_connection():
                    consecutive_errors += 1
                    if consecutive_errors >= max_consecutive_errors:
                        self.log_message("❌ Demasiados errores consecutivos, deteniendo", "error")
                        break
                    if self._wait_with_interrupt(self.retry_delay):
                        break
                    continue

                # Procesar correos
                success = self._process_cargador_emails()

                if success:
                    consecutive_errors = 0
                    self.session_stats['cycles_completed'] += 1
                    self.session_stats['last_check_time'] = datetime.now()
                else:
                    consecutive_errors += 1

                # Esperar intervalo fijo de 1 minuto
                if not self._wait_fixed_interval():
                    break

            except Exception as e:
                self.log_message(f"❌ Error en ciclo de monitoreo: {str(e)}", "error")
                self.session_stats['errors'] += 1
                consecutive_errors += 1

                if consecutive_errors >= max_consecutive_errors:
                    self.log_message("❌ Demasiados errores consecutivos", "error")
                    break

                if self._wait_with_interrupt(self.retry_delay):
                    break

    def _ensure_connection(self):
        """Verifica y restablece conexión si es necesario."""
        if self.stop_event.is_set():
            return False

        if not self.email_manager.is_connected:
            self.log_message("🔄 Conexión perdida, reestableciendo...", "warning")
            return self._establish_connection()

        try:
            if self.stop_event.is_set():
                return False

            status = self.email_manager.connection.noop()
            if status[0] != 'OK':
                self.log_message("🔄 Conexión no responde, reestableciendo...", "warning")
                self.email_manager.disconnect()
                return self._establish_connection()
        except:
            if not self.stop_event.is_set():
                self.log_message("🔄 Error verificando conexión, reestableciendo...", "warning")
                self.email_manager.disconnect()
                return self._establish_connection()
            return False

        return True

    def _process_cargador_emails(self):
        """Procesa correos 'Cargador' con la nueva lógica robusta."""
        if self.stop_event.is_set():
            return False

        try:
            self.log_message("🔍 Buscando correos 'Cargador' con archivos Excel...", "info")

            # NUEVA BÚSQUEDA ROBUSTA
            matching_emails = self.email_manager.search_cargador_emails_with_excel()

            if not matching_emails:
                self.log_message("📭 No se encontraron correos 'Cargador' nuevos", "info")
                return True  # No es error, simplemente no hay correos

            self.log_message(f"📧 Encontrados {len(matching_emails)} correos para procesar", "success")
            self.session_stats['emails_found'] += len(matching_emails)

            # Procesar cada email
            downloaded_files = []
            for email_id in matching_emails:
                if self.stop_event.is_set():
                    break

                try:
                    files = self._process_single_email(email_id)
                    if files:
                        downloaded_files.extend(files)
                except Exception as e:
                    self.log_message(f"❌ Error procesando email: {str(e)}", "error")
                    self.session_stats['errors'] += 1

            if not downloaded_files:
                self.log_message("📭 No se descargaron archivos", "info")
                return True

            self.session_stats['files_downloaded'] += len(downloaded_files)
            self.log_message(f"💾 Total archivos descargados: {len(downloaded_files)}", "success")

            if self.stop_event.is_set():
                return False

            # Procesar con XML si está configurado
            processed_files = downloaded_files
            if self.xml_config and self.excel_processor:
                processing_result = self._process_with_xml(downloaded_files)
                if processing_result and processing_result.get('processed_files'):
                    processed_files = processing_result['processed_files']

            if self.stop_event.is_set():
                return False

            # Enviar archivos si está configurado
            if self.recipients_config and self.email_sender and processed_files:
                self._send_processed_files(processed_files)

            return True

        except Exception as e:
            self.log_message(f"❌ Error en procesamiento: {str(e)}", "error")
            self.session_stats['errors'] += 1
            return False

    def _process_single_email(self, email_id):
        """Procesa un email individual."""
        if self.stop_event.is_set():
            return None

        email_details = self.email_manager.get_email_details(email_id)
        if not email_details:
            return None

        if not email_details['has_excel']:
            # Marcar como leído aunque no tenga Excel
            self.email_manager.mark_email_as_read_and_cache(email_id)
            return None

        # Descargar archivos Excel
        downloaded_files = self.email_manager.download_excel_attachments(
            email_details, self.download_folder
        )

        if downloaded_files:
            for file_info in downloaded_files:
                self.log_message(f"💾 Descargado: {file_info['filename']}", "success")

            # IMPORTANTE: Marcar como leído Y cachear DESPUÉS del procesamiento exitoso
            if self.email_manager.mark_email_as_read_and_cache(email_id):
                self.log_message("✅ Email procesado y marcado como leído", "success")

            return downloaded_files

        return None

    def _process_with_xml(self, downloaded_files):
        """Procesa archivos con XML si está configurado."""
        if not self.excel_processor or not downloaded_files:
            return None

        try:
            self.log_message("🔄 Procesando con datos XML...", "info")

            result = self.excel_processor.process_excel_files(
                input_folder=self.download_folder,
                xml_folder=None,
                output_folder=self.xml_config['output_folder'],
                config=self.xml_config
            )

            if result['success']:
                processed_files = result.get('processed_files', [])
                if processed_files:
                    self.session_stats['files_processed'] += len(processed_files)
                    self.log_message(f"✅ Procesamiento XML completado: {len(processed_files)} archivos", "success")
                return result

        except Exception as e:
            self.log_message(f"❌ Error en procesamiento XML: {str(e)}", "error")

        return None

    def _send_processed_files(self, processed_files):
        """Envía archivos procesados si está configurado."""
        if not self.email_sender or not processed_files:
            return

        try:
            if not self.xml_config.get('auto_send', True):
                self.log_message("📧 Envío automático deshabilitado", "info")
                return

            self.log_message(f"📧 Enviando {len(processed_files)} archivos procesados...", "info")

            result = self.email_sender.send_processed_files(processed_files=processed_files)

            if result['success']:
                self.session_stats['emails_sent'] += result.get('sent_count', 0)
                self.log_message("✅ Envío completado exitosamente", "success")
            else:
                self.log_message(f"❌ Error en envío: {result.get('error', 'Desconocido')}", "error")

        except Exception as e:
            self.log_message(f"❌ Error enviando archivos: {str(e)}", "error")

    def _wait_fixed_interval(self):
        """Espera el intervalo fijo de 1 minuto."""
        interval_seconds = self.MONITORING_INTERVAL_MINUTES * 60
        self.log_message(f"⏰ Esperando {self.MONITORING_INTERVAL_MINUTES} minuto hasta la próxima verificación...",
                         "info")
        return not self._wait_with_interrupt(interval_seconds)

    def _wait_with_interrupt(self, seconds):
        """Espera con verificación de interrupciones."""
        elapsed = 0
        while elapsed < seconds and self.is_running and not self.stop_event.is_set():
            sleep_time = min(self.stop_check_interval, seconds - elapsed)
            time.sleep(sleep_time)
            elapsed += sleep_time
        return self.stop_event.is_set() or not self.is_running

    def _cleanup(self):
        """Limpia recursos y conexiones."""
        try:
            # Detener procesadores
            for processor in [self.excel_processor, self.email_sender]:
                if processor:
                    try:
                        if hasattr(processor, 'stop_processing'):
                            processor.stop_processing()
                        elif hasattr(processor, 'stop_sending'):
                            processor.stop_sending()
                    except:
                        pass

            # Cerrar conexión de email
            if self.email_manager:
                self.email_manager.disconnect()
                if not self.stop_event.is_set():
                    self.log_message("🔌 Conexión cerrada", "info")

            # Log resumen de sesión
            if not self.stop_event.is_set():
                self._log_session_summary()
            else:
                self.log_message("⏹️ Monitoreo detenido por el usuario", "info")

        except Exception as e:
            self.log_message(f"⚠️ Error en limpieza: {str(e)}", "warning")

    def _log_session_summary(self):
        """Log resumen de la sesión."""
        summary = f"""📊 Resumen de sesión:
   • Ciclos completados: {self.session_stats['cycles_completed']}
   • Emails encontrados: {self.session_stats['emails_found']}
   • Archivos descargados: {self.session_stats['files_downloaded']}
   • Archivos procesados: {self.session_stats['files_processed']}
   • Emails enviados: {self.session_stats['emails_sent']}
   • Errores: {self.session_stats['errors']}
   • Reintentos de conexión: {self.session_stats['connection_retries']}"""

        self.log_message(summary, "info")

        if self.session_stats['last_check_time']:
            last_check = self.session_stats['last_check_time'].strftime("%Y-%m-%d %H:%M:%S")
            self.log_message(f"🕐 Última verificación: {last_check}", "info")

    def log_message(self, message, msg_type="info"):
        """Envía mensaje al log de la interfaz."""
        try:
            if self.automation_tab and hasattr(self.automation_tab, 'add_log_message'):
                self.automation_tab.add_log_message(message, msg_type)
        except Exception:
            print(f"[{msg_type.upper()}] {message}")

    # Métodos de compatibilidad (mantenidos pero simplificados)
    def get_session_statistics(self):
        """Obtiene estadísticas de la sesión."""
        return self.session_stats.copy()

    def update_statistics(self, emails_found=None, files_downloaded=None):
        """Método mantenido para compatibilidad."""
        pass