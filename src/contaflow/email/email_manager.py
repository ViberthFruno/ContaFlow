# email_manager.py
"""
Gestor de conexiones de email para ContaFlow con lógica de búsqueda robusta.
Búsqueda simplificada por asunto "Cargador" con archivos Excel sin depender del estado "no leído".
"""
# Archivos relacionados: email_processor.py, config_manager.py

import imaplib
import ssl
import socket
import email
import os
import re
from datetime import date
from email.header import decode_header


class EmailManager:
    """Clase para gestionar conexiones de email con búsqueda robusta sin estado UNSEEN."""

    def __init__(self):
        """Inicializa el gestor de email con configuración robusta."""
        self.providers_config = {
            "Gmail": ("imap.gmail.com", "smtp.gmail.com"),
            "Outlook": ("outlook.office365.com", "smtp-mail.outlook.com"),
            "Yahoo": ("imap.mail.yahoo.com", "smtp.mail.yahoo.com"),
            "Otro": ("", "")
        }
        self.connection = None
        self.is_connected = False
        self.connection_timeout = 30
        self.operation_timeout = 60

        # Cache para evitar procesar el mismo email múltiples veces
        self.processed_emails_cache = set()

    def get_provider_config(self, provider):
        """Obtiene la configuración del proveedor."""
        imap, smtp = self.providers_config.get(provider, ("", ""))
        return {
            "imap_server": imap, "imap_port": 993,
            "smtp_server": smtp, "smtp_port": 587,
            "use_ssl": True
        }

    def test_connection(self, provider, email, password):
        """Prueba la conexión con el servidor."""
        if not re.match(r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$', email):
            return False, "Formato de email inválido"

        config = self.get_provider_config(provider)
        if not config["imap_server"]:
            return False, "Configuración de servidor no disponible"

        success, message = self._test_imap_connection(config, email, password)
        return (True, "Conexión exitosa") if success else (False, f"Error de conexión: {message}")

    def _test_imap_connection(self, config, email, password):
        """Prueba la conexión IMAP."""
        try:
            socket.setdefaulttimeout(self.connection_timeout)
            context = ssl.create_default_context()

            with imaplib.IMAP4_SSL(config["imap_server"], config["imap_port"], ssl_context=context) as imap:
                imap.login(email, password)
                imap.select('INBOX')
                return True, "Conexión IMAP exitosa"

        except imaplib.IMAP4.error as e:
            error_msg = str(e).lower()
            if "authentication failed" in error_msg or "invalid credentials" in error_msg:
                return False, "Credenciales incorrectas"
            elif "too many simultaneous connections" in error_msg:
                return False, "Demasiadas conexiones simultáneas"
            return False, f"Error IMAP: {str(e)}"

        except socket.gaierror:
            return False, "Error de conexión: Servidor no encontrado"
        except socket.timeout:
            return False, "Error de conexión: Timeout"
        except Exception as e:
            return False, f"Error de conexión: {str(e)}"
        finally:
            socket.setdefaulttimeout(None)

    def connect(self, provider, email, password):
        """Establece conexión con el servidor."""
        try:
            if self.is_connected:
                self.disconnect()

            config = self.get_provider_config(provider)
            socket.setdefaulttimeout(self.connection_timeout)

            self.connection = imaplib.IMAP4_SSL(
                config["imap_server"], config["imap_port"],
                ssl_context=ssl.create_default_context()
            )
            self.connection.login(email, password)
            self.connection.select('INBOX')
            self.is_connected = True
            return True

        except Exception as e:
            self.is_connected = False
            if self.connection:
                try:
                    self.connection.close()
                    self.connection.logout()
                except:
                    pass
                self.connection = None
            print(f"Error en connect: {e}")
            return False
        finally:
            socket.setdefaulttimeout(None)

    def disconnect(self):
        """Desconecta del servidor."""
        if self.connection and self.is_connected:
            try:
                socket.setdefaulttimeout(10)
                self.connection.close()
                self.connection.logout()
            except:
                pass
            finally:
                self.connection = None
                self.is_connected = False
                socket.setdefaulttimeout(None)

    def search_cargador_emails_with_excel(self):
        """
        NUEVA LÓGICA ROBUSTA: Busca correos con asunto 'Cargador' y archivos Excel.
        No depende del estado UNSEEN para ser más robusto.

        Returns:
            List[str]: Lista de IDs de emails que cumplen los criterios
        """
        if not self.is_connected or not self.connection:
            print("Error: No hay conexión IMAP activa")
            return []

        try:
            socket.setdefaulttimeout(self.operation_timeout)

            # Buscar correos de hoy con asunto que contenga "Cargador"
            today_str = date.today().strftime("%d-%b-%Y")
            search_query = f'SINCE "{today_str}" SUBJECT "Cargador"'

            print(f"🔍 Búsqueda robusta: {search_query}")

            status, messages = self.connection.search(None, search_query)
            if status != 'OK':
                print(f"Error en búsqueda IMAP: {status}")
                return []

            message_ids = messages[0].split()
            print(f"📧 Correos encontrados con 'Cargador': {len(message_ids)}")

            if not message_ids:
                return []

            # Filtrar por archivos Excel Y que no hayan sido procesados recientemente
            excel_emails = self._filter_by_excel_and_cache(message_ids)
            print(f"📊 Correos con Excel no procesados: {len(excel_emails)}")

            return excel_emails

        except Exception as e:
            print(f"Error en búsqueda robusta: {e}")
            return []
        finally:
            socket.setdefaulttimeout(None)

    def _filter_by_excel_and_cache(self, message_ids):
        """
        Filtra mensajes que tengan archivos Excel y no estén en cache.

        Args:
            message_ids: Lista de IDs de mensajes

        Returns:
            List[str]: IDs de emails con Excel no procesados
        """
        if not message_ids:
            return []

        filtered_ids = []

        for msg_id in message_ids:
            try:
                # Verificar si ya fue procesado recientemente
                if msg_id.decode() in self.processed_emails_cache:
                    print(f"📝 Email {msg_id.decode()} ya procesado recientemente, omitiendo")
                    continue

                socket.setdefaulttimeout(30)
                email_details = self.get_email_details(msg_id)

                if email_details and email_details.get('has_excel', False):
                    filtered_ids.append(msg_id)
                    print(f"✅ Email con Excel encontrado: {email_details['subject'][:50]}...")

            except Exception as e:
                print(f"Error verificando email {msg_id}: {e}")
            finally:
                socket.setdefaulttimeout(None)

        return filtered_ids

    def _decode_header(self, header_value):
        """Decodifica un header de email."""
        try:
            decoded_parts = decode_header(header_value)
            decoded_string = ""

            for part, encoding in decoded_parts:
                if isinstance(part, bytes):
                    decoded_string += part.decode(encoding or 'utf-8')
                else:
                    decoded_string += part

            return decoded_string.strip()
        except Exception as e:
            print(f"Error decodificando header: {e}")
            return header_value.strip()

    def get_email_details(self, message_id):
        """Obtiene detalles de un email."""
        if not self.is_connected or not self.connection:
            return None

        try:
            socket.setdefaulttimeout(self.operation_timeout)
            status, msg_data = self.connection.fetch(message_id, '(RFC822)')

            if status == 'OK':
                email_message = email.message_from_bytes(msg_data[0][1])
                attachments = self._get_attachments_info(email_message)

                return {
                    'message_id': message_id,
                    'subject': self._decode_header(email_message.get('Subject', 'Sin asunto')),
                    'from': self._decode_header(email_message.get('From', 'Desconocido')),
                    'date': email_message.get('Date', 'Fecha desconocida'),
                    'attachments': attachments,
                    'has_excel': any(att['is_excel'] for att in attachments),
                    'email_object': email_message
                }

        except Exception as e:
            print(f"Error getting email details: {e}")
        finally:
            socket.setdefaulttimeout(None)

        return None

    def _get_attachments_info(self, email_message):
        """Extrae información de adjuntos."""
        attachments = []

        try:
            for part in email_message.walk():
                if part.get_content_disposition() == 'attachment':
                    filename = part.get_filename()
                    if filename:
                        filename = self._decode_header(filename)
                        payload = part.get_payload(decode=True)

                        attachments.append({
                            'filename': filename,
                            'is_excel': filename.lower().endswith(('.xlsx', '.xls')),
                            'size': len(payload) if payload else 0,
                            'part_object': part
                        })

        except Exception as e:
            print(f"Error extrayendo adjuntos: {e}")

        return attachments

    def download_excel_attachments(self, email_details, download_folder):
        """Descarga adjuntos Excel."""
        downloaded_files = []
        os.makedirs(download_folder, exist_ok=True)

        for attachment in email_details['attachments']:
            if not attachment['is_excel']:
                continue

            try:
                content = attachment['part_object'].get_payload(decode=True)
                if not content:
                    continue

                base_filename = attachment['filename']
                file_path = os.path.join(download_folder, base_filename)

                counter = 1
                while os.path.exists(file_path):
                    name, ext = os.path.splitext(base_filename)
                    file_path = os.path.join(download_folder, f"{name}_{counter}{ext}")
                    counter += 1

                with open(file_path, 'wb') as f:
                    f.write(content)

                downloaded_files.append({
                    'filename': os.path.basename(file_path),
                    'path': file_path,
                    'size': len(content),
                    'from_email': email_details['from'],
                    'email_subject': email_details['subject']
                })

                print(f"📥 Descargado: {os.path.basename(file_path)}")

            except Exception as e:
                print(f"Error downloading {attachment['filename']}: {e}")

        return downloaded_files

    def mark_email_as_read_and_cache(self, message_id):
        """
        Marca un email como leído Y lo agrega al cache para evitar reprocesamiento.

        Args:
            message_id: ID del mensaje a marcar

        Returns:
            bool: True si se marcó exitosamente
        """
        if not self.is_connected or not self.connection:
            return False

        try:
            socket.setdefaulttimeout(30)

            # Marcar como leído
            self.connection.store(message_id, '+FLAGS', '\\Seen')

            # Agregar al cache para evitar reprocesamiento
            msg_id_str = message_id.decode() if isinstance(message_id, bytes) else str(message_id)
            self.processed_emails_cache.add(msg_id_str)

            print(f"✅ Email marcado como leído y cacheado: {msg_id_str}")

            # Limpiar cache si se hace muy grande (mantener últimos 100)
            if len(self.processed_emails_cache) > 100:
                # Convertir a lista, tomar los últimos 100, convertir de vuelta a set
                cache_list = list(self.processed_emails_cache)
                self.processed_emails_cache = set(cache_list[-100:])
                print("🧹 Cache de emails limpiado (mantenidos últimos 100)")

            return True

        except Exception as e:
            print(f"Error marking email as read: {e}")
            return False
        finally:
            socket.setdefaulttimeout(None)

    def clear_processed_cache(self):
        """Limpia el cache de emails procesados."""
        self.processed_emails_cache.clear()
        print("🗑️ Cache de emails procesados limpiado")

    def get_processed_cache_size(self):
        """Obtiene el tamaño actual del cache."""
        return len(self.processed_emails_cache)

    def get_connection_status(self):
        """Obtiene el estado de la conexión con información del cache."""
        return {
            'connected': self.is_connected,
            'connection_object': self.connection is not None,
            'cache_size': len(self.processed_emails_cache)
        }

    # MÉTODOS OBSOLETOS - Mantenidos para compatibilidad pero marcados como deprecated
    def search_emails_by_criteria(self, search_criteria):
        """
        OBSOLETO: Método mantenido para compatibilidad.
        Usar search_cargador_emails_with_excel() en su lugar.
        """
        print("⚠️ ADVERTENCIA: search_emails_by_criteria está obsoleto")
        print("💡 Usar search_cargador_emails_with_excel() en su lugar")
        return self.search_cargador_emails_with_excel()

    def mark_email_as_read(self, message_id):
        """
        OBSOLETO: Método mantenido para compatibilidad.
        Usar mark_email_as_read_and_cache() en su lugar.
        """
        print("⚠️ ADVERTENCIA: mark_email_as_read está obsoleto")
        print("💡 Usar mark_email_as_read_and_cache() en su lugar")
        return self.mark_email_as_read_and_cache(message_id)

    def __del__(self):
        """Destructor para cerrar conexión."""
        self.disconnect()