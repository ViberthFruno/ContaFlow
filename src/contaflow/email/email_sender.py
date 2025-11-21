# email_sender.py
"""
Sistema de envío de emails con archivos adjuntos para ContaFlow con información detallada y formato mejorado.
Maneja el envío automático consolidado de archivos Excel procesados con mensajes detallados,
información completa de archivos (matches, revisiones manuales, tamaños) y envío como texto plano
para evitar problemas de formato y color.
"""
# Archivos relacionados: config_manager.py, pdf_generator.py, email_processor.py

import smtplib
import os
import threading
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime
from typing import List, Dict, Optional, Tuple
from config.config_manager import ConfigManager


class EmailSender:
    """Clase para gestionar el envío consolidado de emails con múltiples archivos adjuntos y mensajes detallados."""

    def __init__(self, automation_tab=None):
        """
        Inicializa el sistema de envío de emails.

        Args:
            automation_tab: Referencia a la pestaña de automatización para logging
        """
        self.automation_tab = automation_tab
        self.config_manager = ConfigManager()

        # Configuración SMTP por proveedor
        self.smtp_config = {
            "Gmail": {
                "server": "smtp.gmail.com",
                "port": 587,
                "use_tls": True
            },
            "Outlook": {
                "server": "smtp-mail.outlook.com",
                "port": 587,
                "use_tls": True
            },
            "Yahoo": {
                "server": "smtp.mail.yahoo.com",
                "port": 587,
                "use_tls": True
            }
        }

        # Configuración de límites y timeouts
        self.connection_timeout = 30
        self.send_timeout = 60
        self.max_attachment_size = 25 * 1024 * 1024  # 25MB por archivo individual
        self.max_total_size = 50 * 1024 * 1024  # 50MB total para todos los adjuntos
        self.retry_attempts = 3
        self.retry_delay = 5

        # Control de estado
        self.stop_event = threading.Event()

        # Estadísticas
        self.stats = {
            'emails_sent': 0,
            'emails_failed': 0,
            'attachments_sent': 0,
            'total_size_sent': 0,
            'recipients_reached': 0,
            'send_time': 0,
            'pdf_summary_generated': 0
        }

    def send_processed_files(self, processed_files: List[Dict],
                             custom_message: Optional[str] = None,
                             processing_stats: Optional[Dict] = None,
                             excluded_xmls: Optional[List[Dict]] = None) -> Dict:
        """
        Envía todos los archivos procesados en UN SOLO correo consolidado con información detallada.

        Args:
            processed_files (List[Dict]): Lista de archivos procesados
            custom_message (str): Mensaje personalizado opcional (OBSOLETO - se ignora)
            processing_stats (Dict): Estadísticas del procesamiento incluyendo filtrado por fecha
            excluded_xmls (List[Dict]): Lista de XMLs excluidos por fecha

        Returns:
            Dict: Resultado del envío consolidado
        """
        start_time = time.time()
        self.log_message("📧 Iniciando envío consolidado con información detallada", "info")

        try:
            # Resetear estadísticas
            self._reset_stats()

            # Cargar configuraciones
            email_config, recipients_config = self._load_email_configs()
            if not email_config or not recipients_config:
                return {'success': False, 'error': 'Configuración de email o destinatarios incompleta'}

            # Validar archivos
            if not processed_files:
                return {'success': False, 'error': 'No hay archivos para enviar'}

            # Verificar si se debe detener
            if self.stop_event.is_set():
                return {'success': False, 'error': 'Envío cancelado'}

            # Preparar lista de destinatarios
            recipients = self._prepare_recipients_list(recipients_config)
            if not recipients:
                return {'success': False, 'error': 'No hay destinatarios válidos'}

            # Validar archivos y tamaños
            valid_files = self._validate_files_for_consolidated_send(processed_files)
            if not valid_files:
                return {'success': False, 'error': 'Ningún archivo es válido para envío'}

            if self.stop_event.is_set():
                return {'success': False, 'error': 'Envío cancelado'}

            # Generar PDF de resumen (siempre habilitado ahora)
            pdf_file_path = None
            if self.config_manager.is_pdf_summary_enabled():
                pdf_file_path = self._generate_summary_pdf(valid_files, processing_stats, excluded_xmls)
                if pdf_file_path:
                    self.stats['pdf_summary_generated'] = 1

            # Enviar correo consolidado con información detallada
            send_result = self._send_consolidated_email_with_detailed_info(
                valid_files, email_config, recipients, pdf_file_path, processing_stats, excluded_xmls
            )

            # Limpiar PDF temporal
            if pdf_file_path and os.path.exists(pdf_file_path):
                try:
                    os.remove(pdf_file_path)
                except:
                    pass

            # Calcular tiempo total
            self.stats['send_time'] = time.time() - start_time

            if send_result['success']:
                self.stats['emails_sent'] = 1
                self.stats['attachments_sent'] = len(valid_files)
                if pdf_file_path:
                    self.stats['attachments_sent'] += 1  # Contar el PDF
                self.stats['total_size_sent'] = sum(file_info.get('size', 0) for file_info in valid_files)
                self.stats['recipients_reached'] = len(recipients['all'])

                attachments_total = len(valid_files) + (1 if pdf_file_path else 0)
                self.log_message(f"✅ Correo consolidado enviado exitosamente con {attachments_total} archivos adjuntos",
                                 "success")
                self._log_consolidated_sending_summary(attachments_total)
            else:
                self.stats['emails_failed'] = 1
                self.log_message(f"❌ Error enviando correo consolidado: {send_result['error']}", "error")

            return {
                'success': send_result['success'],
                'sent_count': 1 if send_result['success'] else 0,
                'total_count': 1,
                'attachments_count': len(valid_files) + (1 if pdf_file_path else 0) if send_result['success'] else 0,
                'stats': self.stats.copy(),
                'stopped_by_user': self.stop_event.is_set(),
                'error': send_result.get('error')
            }

        except Exception as e:
            self.log_message(f"❌ Error crítico en envío consolidado: {str(e)}", "error")
            return {'success': False, 'error': str(e)}

    def _generate_summary_pdf(self, valid_files: List[Dict], processing_stats: Optional[Dict] = None,
                              excluded_xmls: Optional[List[Dict]] = None) -> Optional[str]:
        """
        Genera un PDF con el resumen detallado del procesamiento incluyendo XMLs excluidos.

        Args:
            valid_files (List[Dict]): Lista de archivos procesados
            processing_stats (Dict): Estadísticas del procesamiento
            excluded_xmls (List[Dict]): Lista de XMLs excluidos por fecha

        Returns:
            str: Ruta del archivo PDF generado o None si falló
        """
        try:
            self.log_message("📄 Generando PDF de resumen detallado con información completa", "info")

            # Importar y usar el generador de PDF
            try:
                from processors.pdf_generator import generate_processing_summary_pdf

                # Generar PDF temporal
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                pdf_filename = f"contaflow_resumen_{timestamp}.pdf"
                pdf_path = os.path.join(os.path.expanduser("~"), "Downloads", pdf_filename)

                success = generate_processing_summary_pdf(
                    processed_files=valid_files,
                    output_path=pdf_path,
                    processing_stats=processing_stats,
                    excluded_xmls=excluded_xmls or []
                )

                if success:
                    self.log_message(f"✅ PDF de resumen generado: {pdf_filename}", "success")
                    return pdf_path
                else:
                    self.log_message("❌ Error generando PDF de resumen", "error")
                    return None

            except ImportError:
                self.log_message("⚠️ PDFGenerator no disponible, omitiendo PDF de resumen", "warning")
                return None

        except Exception as e:
            self.log_message(f"❌ Error generando PDF de resumen: {str(e)}", "error")
            return None

    def _send_consolidated_email_with_detailed_info(self, valid_files: List[Dict], email_config: Dict,
                                                    recipients: Dict, pdf_file_path: Optional[str] = None,
                                                    processing_stats: Optional[Dict] = None,
                                                    excluded_xmls: Optional[List[Dict]] = None) -> Dict:
        """
        Envía un correo consolidado usando las plantillas del config_manager con información detallada.

        Args:
            valid_files (List[Dict]): Lista de archivos válidos a enviar
            email_config (Dict): Configuración de email
            recipients (Dict): Destinatarios
            pdf_file_path (str): Ruta del PDF de resumen opcional
            processing_stats (Dict): Estadísticas del procesamiento
            excluded_xmls (List[Dict]): Lista de XMLs excluidos por fecha

        Returns:
            Dict: Resultado del envío
        """
        for attempt in range(self.retry_attempts):
            if self.stop_event.is_set():
                return {'success': False, 'error': 'Envío cancelado'}

            try:
                self.log_message(
                    f"📤 Enviando correo con información detallada (intento {attempt + 1}/{self.retry_attempts})",
                    "info"
                )

                # Crear mensaje usando plantillas del config_manager
                msg = MIMEMultipart()
                msg['From'] = email_config['email']
                msg['To'] = recipients['main']
                if recipients['cc']:
                    msg['Cc'] = ', '.join(recipients['cc'])

                # Preparar datos detallados para reemplazar en las plantillas
                template_data = self._prepare_detailed_template_data(valid_files, processing_stats, excluded_xmls)

                # Asunto fijo del config_manager
                fixed_subject = self.config_manager.get_email_subject_template()
                msg['Subject'] = fixed_subject

                # Cuerpo detallado del config_manager con variables reemplazadas
                body_template = self.config_manager.get_email_body_template()
                detailed_body = self._replace_template_variables(body_template, template_data)

                # IMPORTANTE: Asegurar envío como texto plano puro
                msg.attach(MIMEText(detailed_body, 'plain', 'utf-8'))

                # Adjuntar todos los archivos Excel procesados
                for i, file_info in enumerate(valid_files, 1):
                    if self.stop_event.is_set():
                        return {'success': False, 'error': 'Envío cancelado'}

                    file_path = file_info.get('output') or file_info.get('path')
                    filename = file_info['filename']

                    self.log_message(f"📎 Adjuntando archivo {i}/{len(valid_files)}: {filename}", "info")

                    with open(file_path, "rb") as attachment:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(attachment.read())

                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename= {filename}'
                    )
                    msg.attach(part)

                # Adjuntar PDF de resumen si está disponible
                if pdf_file_path and os.path.exists(pdf_file_path):
                    self.log_message("📎 Adjuntando PDF de resumen detallado", "info")

                    with open(pdf_file_path, "rb") as pdf_attachment:
                        pdf_part = MIMEBase('application', 'octet-stream')
                        pdf_part.set_payload(pdf_attachment.read())

                    encoders.encode_base64(pdf_part)
                    pdf_part.add_header(
                        'Content-Disposition',
                        f'attachment; filename= {os.path.basename(pdf_file_path)}'
                    )
                    msg.attach(pdf_part)

                # Obtener configuración SMTP
                smtp_cfg = self._get_smtp_config(email_config['provider'])
                if not smtp_cfg:
                    return {'success': False, 'error': f'Proveedor no soportado: {email_config["provider"]}'}

                # Establecer conexión SMTP y enviar
                self.log_message("🔗 Estableciendo conexión SMTP...", "info")
                server = smtplib.SMTP(smtp_cfg['server'], smtp_cfg['port'], timeout=self.connection_timeout)

                if smtp_cfg['use_tls']:
                    server.starttls()

                server.login(email_config['email'], email_config['password'])

                # Enviar email consolidado como texto plano
                all_recipients = [recipients['main']] + recipients['cc']
                server.sendmail(email_config['email'], all_recipients, msg.as_string())
                server.quit()

                self.log_message("✅ Correo consolidado con información detallada enviado exitosamente", "success")
                return {'success': True, 'recipients': len(all_recipients)}

            except smtplib.SMTPAuthenticationError as e:
                error_msg = "Credenciales de email incorrectas"
                self.log_message(f"❌ {error_msg}: {str(e)}", "error")
                return {'success': False, 'error': error_msg}

            except smtplib.SMTPRecipientsRefused as e:
                error_msg = f"Destinatarios rechazados: {str(e)}"
                self.log_message(f"❌ {error_msg}", "error")
                return {'success': False, 'error': error_msg}

            except smtplib.SMTPServerDisconnected as e:
                if attempt < self.retry_attempts - 1:
                    self.log_message(f"⚠️ Conexión perdida, reintentando en {self.retry_delay}s...", "warning")
                    if not self._wait_with_interrupt(self.retry_delay):
                        continue
                    else:
                        return {'success': False, 'error': 'Envío cancelado'}
                else:
                    error_msg = f"Conexión SMTP perdida: {str(e)}"
                    self.log_message(f"❌ {error_msg}", "error")
                    return {'success': False, 'error': error_msg}

            except Exception as e:
                if attempt < self.retry_attempts - 1:
                    self.log_message(f"⚠️ Error en intento {attempt + 1}: {str(e)}, reintentando...", "warning")
                    if not self._wait_with_interrupt(self.retry_delay):
                        continue
                    else:
                        return {'success': False, 'error': 'Envío cancelado'}
                else:
                    error_msg = f"Error enviando email consolidado: {str(e)}"
                    self.log_message(f"❌ {error_msg}", "error")
                    return {'success': False, 'error': error_msg}

        return {'success': False, 'error': 'Máximo de reintentos alcanzado'}

    def _prepare_detailed_template_data(self, valid_files: List[Dict], processing_stats: Optional[Dict] = None,
                                        excluded_xmls: Optional[List[Dict]] = None) -> Dict:
        """
        Prepara los datos detallados para reemplazar las variables en las plantillas de mensaje.

        Args:
            valid_files (List[Dict]): Lista de archivos válidos
            processing_stats (Dict): Estadísticas del procesamiento
            excluded_xmls (List[Dict]): Lista de XMLs excluidos por fecha

        Returns:
            Dict: Diccionario con los datos detallados para las plantillas
        """
        file_count = len(valid_files)
        total_matches = sum(file_info.get('matches', 0) for file_info in valid_files)

        # Crear detalle completo de archivos con TODA la información
        detalle_archivos_completo = ""
        for i, file_info in enumerate(valid_files, 1):
            company_name = file_info.get('company_name', 'Empresa desconocida')
            filename = file_info.get('filename', 'archivo_procesado.xlsx')
            matches = file_info.get('matches', 0)
            manual_reviews = file_info.get('manual_reviews', 0)

            # Calcular tamaño del archivo en MB
            file_size_bytes = file_info.get('size', 0)
            file_size_mb = file_size_bytes / (1024 * 1024) if file_size_bytes > 0 else 0

            # Formato detallado con toda la información requerida
            detalle_archivos_completo += f"{i}. {filename}\n"
            detalle_archivos_completo += f"   • Empresa: {company_name}\n"
            detalle_archivos_completo += f"   • Matches encontrados: {matches:,}\n"
            detalle_archivos_completo += f"   • Revisiones manuales: {manual_reviews}\n"
            detalle_archivos_completo += f"   • Tamaño: {file_size_mb:.2f} MB\n"

            # Agregar línea en blanco entre archivos (excepto el último)
            if i < len(valid_files):
                detalle_archivos_completo += "\n"

        return {
            'total_archivos': file_count,
            's_plural': 's' if file_count > 1 else '',
            'matches_totales': total_matches,
            'detalle_archivos_completo': detalle_archivos_completo.strip(),
            'fecha_procesamiento': datetime.now().strftime("%d/%m/%Y a las %H:%M")
        }

    def _replace_template_variables(self, template: str, data: Dict) -> str:
        """
        Reemplaza las variables en una plantilla con los datos reales.

        Args:
            template (str): Plantilla con variables
            data (Dict): Datos para reemplazar

        Returns:
            str: Plantilla con variables reemplazadas
        """
        try:
            result = template
            for key, value in data.items():
                result = result.replace(f"{{{key}}}", str(value))
            return result
        except Exception as e:
            self.log_message(f"⚠️ Error reemplazando variables en plantilla: {str(e)}", "warning")
            return template

    def stop_sending(self):
        """Detiene el envío de emails en curso."""
        self.stop_event.set()
        self.log_message("⏹️ Señal de parada enviada al sistema de envío", "warning")

    def _reset_stats(self):
        """Resetea las estadísticas para un nuevo envío."""
        self.stats = {
            'emails_sent': 0,
            'emails_failed': 0,
            'attachments_sent': 0,
            'total_size_sent': 0,
            'recipients_reached': 0,
            'send_time': 0,
            'pdf_summary_generated': 0
        }

    def _load_email_configs(self) -> Tuple[Optional[Dict], Optional[Dict]]:
        """
        Carga las configuraciones de email y destinatarios.

        Returns:
            Tuple[Dict, Dict]: Configuración de email y destinatarios
        """
        try:
            config = self.config_manager.load_config()
            if not config:
                self.log_message("❌ No se encontró configuración", "error")
                return None, None

            # Validar configuración de email
            email_fields = ['provider', 'email', 'password']
            for field in email_fields:
                if not config.get(field):
                    self.log_message(f"❌ Campo de email faltante: {field}", "error")
                    return None, None

            email_config = {
                'provider': config['provider'],
                'email': config['email'],
                'password': config['password']
            }

            # Validar configuración de destinatarios
            recipients_config = config.get('recipients_config')
            if not recipients_config or not recipients_config.get('main_recipient'):
                self.log_message("❌ Configuración de destinatarios incompleta", "error")
                return None, None

            self.log_message("✅ Configuraciones cargadas correctamente", "success")
            return email_config, recipients_config

        except Exception as e:
            self.log_message(f"❌ Error cargando configuraciones: {str(e)}", "error")
            return None, None

    def _prepare_recipients_list(self, recipients_config: Dict) -> Optional[Dict]:
        """
        Prepara la lista de destinatarios para envío.

        Args:
            recipients_config (Dict): Configuración de destinatarios

        Returns:
            Dict: Diccionario con listas de destinatarios
        """
        try:
            main_recipient = recipients_config['main_recipient'].strip()
            cc_recipients = [cc.strip() for cc in recipients_config.get('cc_recipients', []) if cc.strip()]

            # Validar formato de emails
            all_emails = [main_recipient] + cc_recipients
            valid_emails = []

            for email in all_emails:
                if self._validate_email_format(email):
                    valid_emails.append(email)
                else:
                    self.log_message(f"⚠️ Email inválido omitido: {email}", "warning")

            if not valid_emails:
                self.log_message("❌ No hay emails válidos en los destinatarios", "error")
                return None

            recipients = {
                'main': main_recipient if main_recipient in valid_emails else valid_emails[0],
                'cc': [email for email in cc_recipients if email in valid_emails],
                'all': valid_emails
            }

            self.log_message(f"👥 Destinatarios preparados: 1 principal + {len(recipients['cc'])} CC", "info")
            return recipients

        except Exception as e:
            self.log_message(f"❌ Error preparando destinatarios: {str(e)}", "error")
            return None

    def _validate_files_for_consolidated_send(self, processed_files: List[Dict]) -> List[Dict]:
        """
        Valida archivos para envío consolidado considerando límites individuales y totales.

        Args:
            processed_files (List[Dict]): Lista de archivos procesados

        Returns:
            List[Dict]: Lista de archivos válidos con información de tamaño
        """
        valid_files = []
        total_size = 0

        self.log_message(f"📋 Validando {len(processed_files)} archivos para envío consolidado", "info")

        for i, file_info in enumerate(processed_files, 1):
            try:
                file_path = file_info.get('output') or file_info.get('path')
                if not file_path or not os.path.exists(file_path):
                    self.log_message(f"⚠️ Archivo {i} no encontrado: {file_path}", "warning")
                    continue

                file_size = os.path.getsize(file_path)

                # Verificar tamaño individual del archivo
                if file_size > self.max_attachment_size:
                    size_mb = file_size / (1024 * 1024)
                    max_mb = self.max_attachment_size / (1024 * 1024)
                    self.log_message(
                        f"⚠️ Archivo {i} muy grande ({size_mb:.1f}MB > {max_mb:.1f}MB): {os.path.basename(file_path)}",
                        "warning"
                    )
                    continue

                # Verificar que el total no exceda el límite
                if total_size + file_size > self.max_total_size:
                    total_mb = (total_size + file_size) / (1024 * 1024)
                    max_total_mb = self.max_total_size / (1024 * 1024)
                    self.log_message(
                        f"⚠️ Agregar archivo {i} excedería límite total ({total_mb:.1f}MB > {max_total_mb:.1f}MB)",
                        "warning"
                    )
                    break  # Parar aquí para no exceder límite total

                # Archivo válido - agregar información de tamaño
                file_info_copy = file_info.copy()
                file_info_copy['size'] = file_size
                file_info_copy['size_mb'] = file_size / (1024 * 1024)
                file_info_copy['filename'] = file_info_copy.get('filename') or os.path.basename(file_path)

                valid_files.append(file_info_copy)
                total_size += file_size

                self.log_message(
                    f"✅ Archivo {i} válido: {file_info_copy['filename']} ({file_info_copy['size_mb']:.2f}MB)",
                    "info"
                )

            except Exception as e:
                self.log_message(f"❌ Error validando archivo {i}: {str(e)}", "error")

        total_mb = total_size / (1024 * 1024)
        self.log_message(
            f"📊 Validación completada: {len(valid_files)}/{len(processed_files)} archivos válidos ({total_mb:.2f}MB total)",
            "success" if valid_files else "warning"
        )

        return valid_files

    def _get_smtp_config(self, provider: str) -> Optional[Dict]:
        """
        Obtiene la configuración SMTP para el proveedor especificado.

        Args:
            provider (str): Nombre del proveedor

        Returns:
            Dict: Configuración SMTP o None si no se encuentra
        """
        return self.smtp_config.get(provider)

    def _validate_email_format(self, email: str) -> bool:
        """
        Valida el formato básico de un email.

        Args:
            email (str): Email a validar

        Returns:
            bool: True si el formato es válido
        """
        import re
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return bool(re.match(pattern, email.strip()))

    def _wait_with_interrupt(self, seconds: float) -> bool:
        """
        Espera con verificación de interrupción.

        Args:
            seconds (float): Segundos a esperar

        Returns:
            bool: True si fue interrumpido, False si terminó normalmente
        """
        elapsed = 0
        check_interval = 0.1

        while elapsed < seconds and not self.stop_event.is_set():
            sleep_time = min(check_interval, seconds - elapsed)
            time.sleep(sleep_time)
            elapsed += sleep_time

        return self.stop_event.is_set()

    def _log_consolidated_sending_summary(self, attachment_count: int):
        """
        Registra un resumen del envío consolidado.

        Args:
            attachment_count (int): Número de archivos adjuntos (incluyendo PDF)
        """
        self.log_message("=" * 60, "info")
        self.log_message("📧 RESUMEN DE ENVÍO CONSOLIDADO CON INFORMACIÓN DETALLADA", "info")
        self.log_message("=" * 60, "info")

        self.log_message(f"✅ Correo consolidado enviado: 1", "info")
        self.log_message(f"📎 Archivos adjuntos: {attachment_count}", "info")
        self.log_message(f"👥 Destinatarios alcanzados: {self.stats['recipients_reached']}", "info")

        # Mostrar si se generó PDF
        if self.stats['pdf_summary_generated'] > 0:
            self.log_message(f"📄 PDF de resumen incluido: ✅", "info")
        else:
            self.log_message(f"📄 PDF de resumen: ❌ (error en generación)", "info")

        # Tamaño total enviado
        if self.stats['total_size_sent'] > 0:
            size_mb = self.stats['total_size_sent'] / (1024 * 1024)
            self.log_message(f"📊 Datos enviados: {size_mb:.2f} MB", "info")

        # Tiempo de envío
        send_time = self.stats['send_time']
        if send_time > 60:
            time_str = f"{send_time / 60:.1f} minutos"
        else:
            time_str = f"{send_time:.1f} segundos"
        self.log_message(f"⏱️ Tiempo total de envío: {time_str}", "info")

        self.log_message("📈 Envío consolidado con información detallada completado exitosamente", "success")
        self.log_message("=" * 60, "info")

    def test_email_connection(self, email_config: Dict) -> Tuple[bool, str]:
        """
        Prueba la conexión SMTP con las credenciales proporcionadas.

        Args:
            email_config (Dict): Configuración de email

        Returns:
            Tuple[bool, str]: (éxito, mensaje)
        """
        try:
            smtp_cfg = self._get_smtp_config(email_config['provider'])
            if not smtp_cfg:
                return False, f"Proveedor no soportado: {email_config['provider']}"

            self.log_message("🔌 Probando conexión SMTP...", "info")

            server = smtplib.SMTP(smtp_cfg['server'], smtp_cfg['port'], timeout=self.connection_timeout)

            if smtp_cfg['use_tls']:
                server.starttls()

            server.login(email_config['email'], email_config['password'])
            server.quit()

            self.log_message("✅ Conexión SMTP exitosa", "success")
            return True, "Conexión exitosa"

        except smtplib.SMTPAuthenticationError:
            error_msg = "Credenciales incorrectas"
            self.log_message(f"❌ {error_msg}", "error")
            return False, error_msg

        except Exception as e:
            error_msg = f"Error de conexión: {str(e)}"
            self.log_message(f"❌ {error_msg}", "error")
            return False, error_msg

    def log_message(self, message: str, msg_type: str = "info"):
        """
        Envía un mensaje al log de la interfaz de forma segura.

        Args:
            message (str): Mensaje a registrar
            msg_type (str): Tipo de mensaje
        """
        try:
            if self.automation_tab and hasattr(self.automation_tab, 'add_log_message'):
                self.automation_tab.add_log_message(message, msg_type)
            else:
                # Fallback a print si no hay automation_tab
                print(f"[{msg_type.upper()}] {message}")
        except Exception as e:
            # Último recurso si hay problemas con logging
            print(f"[{msg_type.upper()}] {message}")
            print(f"[LOG ERROR] {e}")

    def get_sending_stats(self) -> Dict:
        """
        Obtiene las estadísticas actuales del envío.

        Returns:
            Dict: Copia de las estadísticas actuales
        """
        return self.stats.copy()