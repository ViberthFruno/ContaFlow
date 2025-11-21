# config_manager.py
"""
Gestor de configuraciones simplificado para ContaFlow con rutas dinámicas XML.
Maneja almacenamiento y recuperación segura de configuraciones sin auto-inicio del bot.
"""
# Archivos relacionados: automatizacion_tab.py, configuracion_tab.py, xml_tab.py, busqueda_tab.py, email_sender.py

import json
import os
import base64
import re
import threading
from datetime import datetime


class ConfigManager:
    """Clase para gestionar las configuraciones de la aplicación con rutas dinámicas XML (sin auto-inicio)."""

    def __init__(self):
        # 🔧 MEJORADO: Usar directorio de configuración en home del usuario
        self.config_dir = self._get_config_directory()
        self.config_file = os.path.join(self.config_dir, "contaflow_config.json")
        self._lock = threading.Lock()
        self._ensure_files_exist()

    def _get_config_directory(self):
        """🆕 Obtiene o crea el directorio de configuración de ContaFlow.

        Returns:
            str: Ruta al directorio de configuración
        """
        # Intentar usar el directorio actual primero
        current_dir = os.getcwd()
        config_dir = os.path.join(current_dir, "config")

        # Si no se puede crear en el directorio actual, usar el home del usuario
        try:
            os.makedirs(config_dir, exist_ok=True)
            # Verificar que podemos escribir
            test_file = os.path.join(config_dir, ".test")
            with open(test_file, 'w') as f:
                f.write("test")
            os.remove(test_file)
            return config_dir
        except (PermissionError, OSError):
            # Fallback: usar directorio home del usuario
            home_dir = os.path.expanduser("~")
            config_dir = os.path.join(home_dir, ".contaflow")
            os.makedirs(config_dir, exist_ok=True)
            return config_dir

    def _ensure_files_exist(self):
        """Asegura que el directorio y archivo de configuración existan."""
        try:
            # 🔧 MEJORADO: Asegurar que el directorio existe
            os.makedirs(self.config_dir, exist_ok=True)

            if not os.path.exists(self.config_file):
                self._create_empty_config()
        except Exception as e:
            print(f"Error creando archivo de configuración: {e}")

    def _create_empty_config(self):
        """Crea un archivo de configuración vacío con campos por defecto."""
        empty_config = {
            "version": "2.0",
            "system_type": "simplified_cargador_search",
            "created": datetime.now().isoformat(),
            "last_updated": None,
            "monitoring_interval": "1_minute_fixed",
            "features_removed": ["auto_start", "configurable_interval"]
        }
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(empty_config, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Error creando configuración vacía: {e}")

    # ========== MÉTODOS PARA MENSAJES DE EMAIL MEJORADOS ==========

    def get_email_subject_template(self):
        """
        Obtiene la plantilla del asunto del email con timestamp.

        Returns:
            str: Asunto con timestamp para todos los emails
        """
        timestamp = datetime.now().strftime("%d/%m/%Y %H:%M")
        return f"Reporte Automatico de Contabilidad - {timestamp}"

    def get_email_body_template(self):
        """
        Obtiene la plantilla mejorada del cuerpo del email con formato detallado.

        Returns:
            str: Cuerpo del email con variables dinámicas y formato mejorado
        """
        return """Estimado/a,

Se adjuntan {total_archivos} archivo{s_plural} Excel procesado{s_plural} con {matches_totales} registros encontrados correspondientes al procesamiento XML empresarial.

📎 Archivos incluidos:
{detalle_archivos_completo}

📄 El PDF adjunto contiene el resumen detallado del procesamiento con información adicional sobre matches, revisiones manuales y estadísticas por empresa.

Este es un correo generado automáticamente por ContaFlow v2.0 (Sistema Simplificado). Por favor, no responda a este mensaje. En caso de cualquier inconveniente o consulta sobre los archivos procesados, comuníquese con el departamento de TI.

Saludos cordiales,
Sistema ContaFlow v2.0

Procesado el: {fecha_procesamiento}
Sistema: Búsqueda optimizada correos "Cargador" con archivos Excel"""

    def is_pdf_summary_enabled(self):
        """
        Verifica si el envío de PDF de resumen está habilitado.

        Returns:
            bool: Siempre True ya que el PDF siempre se envía
        """
        return True

    # ========== MÉTODOS PARA RUTAS DINÁMICAS XML ==========

    def build_dynamic_xml_path(self, base_path):
        """
        Construye una ruta XML dinámica agregando año/mes actual a la carpeta base.

        Args:
            base_path (str): Ruta base de la carpeta (ej: "V:\\3101263133")

        Returns:
            str: Ruta completa con año/mes (ej: "V:\\3101263133\\2025\\7")
        """
        try:
            if not base_path or not base_path.strip():
                return base_path

            # Limpiar la ruta base
            base_path = base_path.strip()

            # Obtener fecha actual
            current_date = datetime.now()
            year = current_date.year
            month = current_date.month

            # Construir ruta dinámica: base\año\mes
            dynamic_path = os.path.join(base_path, str(year), str(month))

            return dynamic_path

        except Exception as e:
            print(f"Error construyendo ruta dinámica para {base_path}: {e}")
            return base_path  # Devolver ruta original si hay error

    def validate_dynamic_xml_path(self, base_path):
        """
        Valida si existe la ruta XML dinámica para el mes/año actual.

        Args:
            base_path (str): Ruta base de la carpeta

        Returns:
            tuple: (existe, ruta_completa, mensaje)
        """
        try:
            if not base_path or not base_path.strip():
                return False, base_path, "Ruta base vacía"

            # Construir ruta dinámica
            dynamic_path = self.build_dynamic_xml_path(base_path)

            # Verificar si existe
            if os.path.exists(dynamic_path) and os.path.isdir(dynamic_path):
                return True, dynamic_path, f"Carpeta del mes actual encontrada"
            else:
                current_date = datetime.now()
                return False, dynamic_path, f"Carpeta {current_date.month}/{current_date.year} no existe"

        except Exception as e:
            return False, base_path, f"Error validando ruta: {str(e)}"

    def get_all_dynamic_xml_paths(self, company_folders_config):
        """
        Obtiene todas las rutas XML dinámicas para las empresas configuradas.

        Args:
            company_folders_config (dict): Configuración de carpetas por empresa

        Returns:
            dict: {company_key: {'base_path': str, 'dynamic_path': str, 'exists': bool, 'message': str}}
        """
        result = {}

        try:
            if not company_folders_config:
                return result

            for company_key, base_path in company_folders_config.items():
                exists, dynamic_path, message = self.validate_dynamic_xml_path(base_path)

                result[company_key] = {
                    'base_path': base_path,
                    'dynamic_path': dynamic_path,
                    'exists': exists,
                    'message': message
                }

        except Exception as e:
            print(f"Error obteniendo rutas dinámicas: {e}")

        return result

    def get_current_month_folder_info(self):
        """
        Obtiene información sobre la carpeta del mes actual.

        Returns:
            dict: Información del mes/año actual
        """
        current_date = datetime.now()
        return {
            'year': current_date.year,
            'month': current_date.month,
            'folder_suffix': f"{current_date.year}\\{current_date.month}",
            'display_text': f"{current_date.month:02d}/{current_date.year}"
        }

    def is_using_dynamic_paths(self):
        """
        Verifica si la configuración actual usa rutas dinámicas.

        Returns:
            bool: True si está configurado para usar rutas dinámicas
        """
        try:
            config = self.load_config()
            if not config:
                return False

            xml_config = config.get('xml_config', {})

            # Si tiene company_folders, asumimos que usa rutas dinámicas
            return 'company_folders' in xml_config and xml_config['company_folders']

        except Exception:
            return False

    # ========== MÉTODOS PRINCIPALES DE CONFIGURACIÓN ==========

    def _encode_password(self, password):
        """Codifica una contraseña usando base64."""
        try:
            return base64.b64encode(password.encode('utf-8')).decode('utf-8')
        except Exception:
            return password

    def _decode_password(self, encoded_password):
        """Decodifica una contraseña desde base64."""
        try:
            return base64.b64decode(encoded_password.encode('utf-8')).decode('utf-8')
        except Exception:
            return encoded_password

    def save_config(self, config_data):
        """Guarda la configuración en un archivo JSON."""
        with self._lock:
            try:
                is_valid, error_msg = self.validate_config(config_data)
                if not is_valid:
                    raise Exception(f"Configuración inválida: {error_msg}")

                current_config = self._load_raw_config()
                current_config.update(config_data)

                if 'password' in current_config:
                    current_config['password'] = self._encode_password(current_config['password'])

                # Actualizar metadatos del sistema
                current_config['last_updated'] = datetime.now().isoformat()
                current_config['version'] = "2.0"
                current_config['system_type'] = "simplified_cargador_search"

                # 🔧 MEJORADO: Asegurar que el directorio existe antes de guardar
                os.makedirs(self.config_dir, exist_ok=True)

                with open(self.config_file, 'w', encoding='utf-8') as f:
                    json.dump(current_config, f, indent=2, ensure_ascii=False)

                print(f"✅ Configuración guardada en: {self.config_file}")

            except Exception as e:
                raise Exception(f"Error guardando configuración: {str(e)}")

    def load_config(self):
        """Carga la configuración desde el archivo JSON."""
        with self._lock:
            try:
                config_data = self._load_raw_config()
                if not config_data:
                    return None

                if 'password' in config_data:
                    config_data['password'] = self._decode_password(config_data['password'])

                return config_data

            except Exception as e:
                print(f"Error cargando configuración: {e}")
                return None

    def _load_raw_config(self):
        """Carga la configuración sin procesar."""
        try:
            if not os.path.exists(self.config_file):
                return {}

            with open(self.config_file, 'r', encoding='utf-8') as f:
                return json.load(f)

        except (json.JSONDecodeError, FileNotFoundError):
            self._create_empty_config()
            return {}
        except Exception as e:
            print(f"Error cargando configuración raw: {e}")
            return {}

    def update_config(self, updates):
        """Actualiza parcialmente la configuración existente."""
        try:
            current_config = self.load_config() or {}
            current_config.update(updates)
            self.save_config(current_config)
        except Exception as e:
            raise Exception(f"Error actualizando configuración: {str(e)}")

    def clear_config(self):
        """Elimina la configuración guardada."""
        with self._lock:
            try:
                if os.path.exists(self.config_file):
                    os.remove(self.config_file)
                    self._create_empty_config()
            except Exception as e:
                raise Exception(f"Error eliminando configuración: {str(e)}")

    def config_exists(self):
        """Verifica si existe una configuración válida."""
        try:
            config = self.load_config()
            return bool(config and config.get('email') and config.get('password'))
        except Exception:
            return False

    # ========== MÉTODOS DE VALIDACIÓN SIMPLIFICADOS ==========

    def validate_config(self, config_data):
        """Valida que la configuración tenga los campos necesarios (sin auto-inicio)."""
        if not isinstance(config_data, dict):
            return False, "Configuración debe ser un diccionario"

        validators = [
            (any(f in config_data for f in ['email', 'password', 'provider']),
             self._validate_email_config),
            ('search_criteria' in config_data,
             lambda d: self._validate_search_criteria(d['search_criteria'])),
            ('xml_config' in config_data,
             lambda d: self._validate_xml_config(d['xml_config'])),
            ('recipients_config' in config_data,
             lambda d: self._validate_recipients_config(d['recipients_config']))
        ]

        for condition, validator in validators:
            if condition:
                is_valid, error = validator(config_data)
                if not is_valid:
                    return False, error

        return True, "Configuración válida"

    def _validate_email_config(self, config_data):
        """Valida configuración de email."""
        email_fields = ['provider', 'email', 'password']
        for field in email_fields:
            if field not in config_data or not config_data[field]:
                return False, f"Campo de email requerido: {field}"

        if not self._validate_email_format(config_data['email']):
            return False, "Formato de email inválido"

        return True, ""

    def _validate_email_format(self, email):
        """Valida el formato básico del email."""
        pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
        return re.match(pattern, email) is not None

    def _validate_search_criteria(self, search_criteria):
        """Valida los criterios de búsqueda simplificados."""
        if not isinstance(search_criteria, dict):
            return False, "Criterios de búsqueda deben ser un diccionario"

        # Para el sistema simplificado, solo validamos la carpeta de descarga
        download_folder = search_criteria.get('download_folder')
        if not download_folder:
            return False, "Campo de búsqueda requerido: download_folder"

        try:
            os.makedirs(download_folder, exist_ok=True)
        except Exception as e:
            return False, f"Carpeta de descarga inválida: {str(e)}"

        # Verificar que tenga configuración fija (o establecerla si no existe)
        expected_config = {
            "subject": "Cargador",
            "search_type": "Contiene",
            "today_only": True,
            "attachments_only": True,
            "excel_files": True
        }

        for key, expected_value in expected_config.items():
            if key not in search_criteria:
                search_criteria[key] = expected_value
            elif search_criteria[key] != expected_value:
                # Advertir si la configuración no coincide con la esperada
                print(f"⚠️ Configuración fija esperada para {key}: {expected_value}, encontrada: {search_criteria[key]}")

        return True, "Criterios de búsqueda válidos"

    def _validate_xml_config(self, xml_config):
        """Valida la configuración XML empresarial incluyendo actividades comerciales y rutas dinámicas."""
        if not isinstance(xml_config, dict):
            return False, "Configuración XML debe ser un diccionario"

        if 'company_folders' in xml_config:
            company_folders = xml_config['company_folders']

            if not isinstance(company_folders, dict) or not company_folders:
                return False, "Debe configurar al menos una carpeta empresarial"

            valid_keys = ['nargallo', 'ventas_fruno', 'creme_caramel', 'su_laka']
            for key, path in company_folders.items():
                if key not in valid_keys:
                    return False, f"Clave de empresa inválida: {key}"
                if not path or not isinstance(path, str):
                    return False, f"Ruta de carpeta inválida para {key}"

                # Para rutas dinámicas, validamos la carpeta base, no la carpeta mes/año
                if not os.path.exists(path):
                    return False, f"Carpeta base no existe para {key}: {path}"
                if not os.access(path, os.R_OK):
                    return False, f"Carpeta base no accesible para {key}: {path}"

            # Validar actividades comerciales (opcional)
            if 'commercial_activities' in xml_config:
                commercial_activities = xml_config['commercial_activities']

                if not isinstance(commercial_activities, dict):
                    return False, "Las actividades comerciales deben ser un diccionario"

                # Validar que las claves de actividades correspondan a empresas válidas
                for company_key, activity in commercial_activities.items():
                    if company_key not in valid_keys:
                        return False, f"Clave de empresa inválida en actividades comerciales: {company_key}"

                    # Validar formato de actividad comercial (opcional, puede estar vacía)
                    if activity is not None and not isinstance(activity, str):
                        return False, f"Actividad comercial para {company_key} debe ser texto"

                    # Validar longitud razonable si no está vacía
                    if activity and len(activity.strip()) > 200:
                        return False, f"Actividad comercial para {company_key} demasiado larga (máximo 200 caracteres)"

        elif 'xml_folder' in xml_config:
            if not xml_config['xml_folder'] or not os.path.exists(xml_config['xml_folder']):
                return False, "Carpeta XML no válida"
        else:
            return False, "Debe especificar company_folders o xml_folder"

        if not xml_config.get('output_folder'):
            return False, "Campo XML requerido: output_folder"

        try:
            os.makedirs(xml_config['output_folder'], exist_ok=True)
        except Exception as e:
            return False, f"Carpeta de salida inválida: {str(e)}"

        manual_limit = xml_config.get('manual_review_limit', 3)
        try:
            if not 1 <= int(manual_limit) <= 20:
                return False, "Límite de revisión manual debe estar entre 1 y 20"
        except (ValueError, TypeError):
            return False, "Límite de revisión manual debe ser un número entero"

        boolean_fields = ['delete_originals', 'auto_send', 'detailed_logs']
        for field in boolean_fields:
            if field in xml_config and not isinstance(xml_config[field], bool):
                return False, f"Campo {field} debe ser verdadero o falso"

        if 'combustible_exclusions' in xml_config:
            is_valid, error = self._validate_combustible_exclusions(
                xml_config['combustible_exclusions']
            )
            if not is_valid:
                return False, error

        return True, "Configuración XML válida"

    def _validate_combustible_exclusions(self, exclusions_config):
        """Valida la estructura de exclusiones de combustible."""
        if exclusions_config is None:
            return True, ""

        emitter_names = []

        if isinstance(exclusions_config, dict):
            emitter_names = exclusions_config.get('emitter_names', [])
        elif isinstance(exclusions_config, list):
            emitter_names = exclusions_config
        else:
            return False, "Exclusiones de combustible deben ser lista o diccionario"

        if emitter_names is None:
            return True, ""

        if not isinstance(emitter_names, list):
            return False, "La lista de emisores excluidos debe ser una lista"

        for name in emitter_names:
            if not isinstance(name, str):
                return False, "Cada NombreEmisor excluido debe ser texto"
            if not name.strip():
                return False, "Los nombres de emisores excluidos no pueden estar vacíos"

        return True, ""

    def _validate_recipients_config(self, recipients_config):
        """Valida la configuración de destinatarios."""
        if not isinstance(recipients_config, dict):
            return False, "Configuración de destinatarios debe ser un diccionario"

        main_recipient = recipients_config.get('main_recipient', '').strip()
        if not main_recipient or not self._validate_email_format(main_recipient):
            return False, "Destinatario principal inválido o faltante"

        cc_recipients = recipients_config.get('cc_recipients', [])
        if not isinstance(cc_recipients, list):
            return False, "Lista de CC debe ser una lista"

        for i, cc_email in enumerate(cc_recipients):
            if cc_email.strip() and not self._validate_email_format(cc_email):
                return False, f"Formato del CC #{i + 1} es inválido: {cc_email}"

        all_emails = [main_recipient] + [cc for cc in cc_recipients if cc.strip()]
        if len(all_emails) != len(set(email.lower() for email in all_emails)):
            return False, "Hay emails duplicados en la configuración de destinatarios"

        return True, "Configuración de destinatarios válida"

    # ========== MÉTODOS DE INFORMACIÓN Y RESUMEN ==========

    def get_config_summary(self):
        """Obtiene un resumen de la configuración actual con información del sistema simplificado."""
        try:
            config = self.load_config()

            if not config:
                return {
                    'config_exists': False,
                    'provider': 'No configurado',
                    'email': 'No configurado',
                    'has_password': False,
                    'has_search_criteria': False,
                    'has_xml_config': False,
                    'has_recipients_config': False,
                    'system_type': 'simplified_cargador_search',
                    'monitoring_interval': '1_minute_fixed',
                    'uses_dynamic_paths': False
                }

            summary = {
                'config_exists': True,
                'provider': config.get('provider', 'No configurado'),
                'email': config.get('email', 'No configurado'),
                'has_password': bool(config.get('password')),
                'has_search_criteria': 'search_criteria' in config,
                'has_xml_config': 'xml_config' in config,
                'has_recipients_config': 'recipients_config' in config,
                'last_updated': config.get('last_updated', 'Desconocido'),
                'system_type': 'simplified_cargador_search',
                'monitoring_interval': '1_minute_fixed',
                'uses_dynamic_paths': self.is_using_dynamic_paths(),
                'version': config.get('version', '2.0')
            }

            if 'search_criteria' in config:
                sc = config['search_criteria']
                summary.update({
                    'search_subject': 'Cargador (fijo)',  # Ahora es fijo
                    'download_folder': sc.get('download_folder', 'No configurado'),
                    'search_type': 'Búsqueda optimizada correos "Cargador" + Excel'
                })

            if 'xml_config' in config:
                xml_config = config['xml_config']

                if 'company_folders' in xml_config:
                    company_names = {
                        'nargallo': 'Nargallo del Este S.A.',
                        'ventas_fruno': 'Ventas Fruno, S.A.',
                        'creme_caramel': 'Creme Caramel',
                        'su_laka': 'Su Laka'
                    }
                    companies = xml_config['company_folders']
                    configured = [company_names.get(k, k) for k in companies.keys()]

                    # Incluir información sobre actividades comerciales
                    commercial_activities = xml_config.get('commercial_activities', {})
                    activities_configured = len([v for v in commercial_activities.values() if v and v.strip()])

                    # Información sobre rutas dinámicas
                    month_info = self.get_current_month_folder_info()
                    dynamic_paths_info = self.get_all_dynamic_xml_paths(companies)

                    summary.update({
                        'xml_companies_configured': len(companies),
                        'xml_companies_list': configured,
                        'xml_activities_configured': activities_configured,
                        'xml_output_folder': xml_config.get('output_folder', 'No configurado'),
                        'xml_auto_send': xml_config.get('auto_send', False),
                        'xml_delete_originals': xml_config.get('delete_originals', True),
                        'dynamic_paths_info': dynamic_paths_info,
                        'current_month_info': month_info
                    })
                else:
                    summary.update({
                        'xml_folder': xml_config.get('xml_folder', 'No configurado'),
                        'xml_output_folder': xml_config.get('output_folder', 'No configurado'),
                        'xml_auto_send': xml_config.get('auto_send', False),
                        'xml_delete_originals': xml_config.get('delete_originals', True),
                        'xml_activities_configured': 0,
                        'uses_dynamic_paths': False
                    })

            if 'recipients_config' in config:
                rc = config['recipients_config']
                cc_count = len(rc.get('cc_recipients', []))
                summary.update({
                    'main_recipient': rc.get('main_recipient', 'No configurado'),
                    'cc_count': cc_count,
                    'total_recipients': 1 + cc_count
                })

            return summary

        except Exception as e:
            print(f"Error obteniendo resumen: {e}")
            return {
                'config_exists': False,
                'error': str(e),
                'system_type': 'simplified_cargador_search',
                'monitoring_interval': '1_minute_fixed',
                'uses_dynamic_paths': False
            }

    def validate_complete_config(self):
        """Valida que la configuración esté completa para operación automática (sin auto-inicio)."""
        try:
            config = self.load_config()
            if not config:
                return False, ['Configuración básica'], []

            missing = []
            warnings = []

            if not all(config.get(f) for f in ['provider', 'email', 'password']):
                missing.append('Credenciales de email')

            if 'search_criteria' not in config:
                missing.append('Criterios de búsqueda')
            elif not config['search_criteria'].get('download_folder'):
                missing.append('Carpeta de descarga')

            if 'xml_config' not in config:
                warnings.append('Configuración XML no configurada - procesamiento XML deshabilitado')
            else:
                xml = config['xml_config']
                if 'company_folders' in xml:
                    if not xml['company_folders'] or not xml.get('output_folder'):
                        warnings.append('Configuración XML incompleta - procesamiento XML deshabilitado')
                    else:
                        # Verificar rutas dinámicas
                        dynamic_paths_info = self.get_all_dynamic_xml_paths(xml['company_folders'])
                        existing_paths = [info for info in dynamic_paths_info.values() if info['exists']]

                        if not existing_paths:
                            warnings.append(f'Sin carpetas XML del mes actual - se procesarán cuando estén disponibles')
                        elif len(existing_paths) < len(dynamic_paths_info):
                            missing_count = len(dynamic_paths_info) - len(existing_paths)
                            warnings.append(f'{missing_count} empresa(s) sin carpeta del mes actual')

                        # Advertencia opcional sobre actividades comerciales
                        commercial_activities = xml.get('commercial_activities', {})
                        activities_configured = len([v for v in commercial_activities.values() if v and v.strip()])
                        companies_configured = len(xml['company_folders'])

                        if activities_configured == 0:
                            warnings.append('Sin actividades comerciales configuradas - se usarán campos vacíos')
                        elif activities_configured < companies_configured:
                            warnings.append(
                                f'Solo {activities_configured} de {companies_configured} empresas tienen actividad comercial configurada')

                elif 'xml_folder' in xml:
                    if not xml['xml_folder'] or not xml.get('output_folder'):
                        warnings.append('Configuración XML incompleta - procesamiento XML deshabilitado')
                else:
                    warnings.append('Configuración XML incompleta - procesamiento XML deshabilitado')

            if 'recipients_config' not in config or not config.get('recipients_config', {}).get('main_recipient'):
                warnings.append('Destinatarios no configurados - envío automático deshabilitado')

            # Información sobre el nuevo sistema
            if self.is_using_dynamic_paths():
                month_info = self.get_current_month_folder_info()
                warnings.append(f'Sistema simplificado v2.0 - Rutas dinámicas activas para mes {month_info["display_text"]}')
            else:
                warnings.append('Sistema simplificado v2.0 - Monitoreo fijo cada 1 minuto de correos "Cargador"')

            return len(missing) == 0, missing, warnings

        except Exception as e:
            return False, [f'Error validando configuración: {str(e)}'], []

    # ========== MÉTODOS DE UTILIDAD ==========

    def backup_config(self, backup_filename=None):
        """Crea una copia de seguridad de la configuración."""
        with self._lock:
            try:
                if not self.config_exists():
                    raise Exception("No hay configuración para respaldar")

                if not backup_filename:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    backup_filename = f"contaflow_v2_backup_{timestamp}.json"

                import shutil
                shutil.copy2(self.config_file, backup_filename)

            except Exception as e:
                raise Exception(f"Error creando respaldo: {str(e)}")

    def restore_config(self, backup_filename):
        """Restaura la configuración desde un archivo de respaldo."""
        with self._lock:
            try:
                if not os.path.exists(backup_filename):
                    raise Exception("Archivo de respaldo no encontrado")

                with open(backup_filename, 'r', encoding='utf-8') as f:
                    json.load(f)  # Validar JSON

                import shutil
                shutil.copy2(backup_filename, self.config_file)

            except Exception as e:
                raise Exception(f"Error restaurando respaldo: {str(e)}")

    def get_xml_config(self):
        """Obtiene solo la configuración XML."""
        try:
            config = self.load_config()
            return config.get('xml_config') if config else None
        except Exception:
            return None

    def get_recipients_config(self):
        """Obtiene solo la configuración de destinatarios."""
        try:
            config = self.load_config()
            return config.get('recipients_config') if config else None
        except Exception:
            return None

    def is_xml_processing_enabled(self):
        """Verifica si el procesamiento XML está habilitado."""
        try:
            xml_config = self.get_xml_config()
            if not xml_config:
                return False

            if 'company_folders' in xml_config:
                return bool(xml_config['company_folders'] and xml_config.get('output_folder'))
            elif 'xml_folder' in xml_config:
                return bool(xml_config['xml_folder'] and xml_config.get('output_folder'))

            return False
        except Exception:
            return False

    def is_auto_send_enabled(self):
        """Verifica si el envío automático está habilitado."""
        try:
            recipients = self.get_recipients_config()
            if not recipients or not recipients.get('main_recipient'):
                return False

            xml_config = self.get_xml_config()
            return xml_config.get('auto_send', True) if xml_config else True

        except Exception:
            return False

    def get_configured_companies(self):
        """Obtiene la lista de empresas configuradas con sus rutas dinámicas."""
        try:
            xml_config = self.get_xml_config()
            if not xml_config:
                return {}

            if 'company_folders' in xml_config:
                # Para rutas dinámicas, devolver las rutas completas con año/mes
                company_folders = xml_config['company_folders']
                dynamic_paths = {}

                for company_key, base_path in company_folders.items():
                    dynamic_path = self.build_dynamic_xml_path(base_path)
                    dynamic_paths[company_key] = dynamic_path

                return dynamic_paths
            elif 'xml_folder' in xml_config and xml_config['xml_folder']:
                return {'nargallo': xml_config['xml_folder']}

            return {}
        except Exception:
            return {}

    def get_configured_commercial_activities(self):
        """Obtiene las actividades comerciales configuradas."""
        try:
            xml_config = self.get_xml_config()
            if not xml_config:
                return {}

            return xml_config.get('commercial_activities', {})
        except Exception:
            return {}

    def get_system_info(self):
        """
        Obtiene información detallada del sistema simplificado.

        Returns:
            dict: Información completa del sistema
        """
        return {
            'version': '2.0',
            'system_type': 'simplified_cargador_search',
            'monitoring_interval': '1_minute_fixed',
            'search_method': 'Correos con asunto "Cargador" + archivos Excel',
            'cache_system': 'Anti-duplicados habilitado',
            'dynamic_paths': 'Rutas XML automáticas año/mes',
            'features_removed': [
                'Auto-inicio del bot',
                'Configuración de intervalo variable',
                'Búsqueda por estado UNSEEN',
                'Configuración compleja de criterios'
            ],
            'features_added': [
                'Búsqueda robusta sin estado UNSEEN',
                'Sistema de cache anti-duplicados',
                'Intervalo de monitoreo optimizado (1 min)',
                'Configuración simplificada',
                'UI limpia sin opciones innecesarias'
            ],
            'configuration_required': [
                'Credenciales de email',
                'Carpeta de descarga'
            ],
            'configuration_optional': [
                'Procesamiento XML por empresa',
                'Envío automático de resultados',
                'Actividades comerciales por empresa'
            ]
        }