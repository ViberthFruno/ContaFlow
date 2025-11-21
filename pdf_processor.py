# pdf_processor.py
"""
Procesador de PDFs mejorado para ContaFlow con soporte robusto para facturas de Correos de Costa Rica SA.
Extrae números de factura y códigos de guías con múltiples patrones y manejo avanzado de errores.
"""

import os
import re
from typing import List, Optional, Tuple, Dict

try:
    import pdfplumber

    PDFPLUMBER_AVAILABLE = True
except ImportError:
    PDFPLUMBER_AVAILABLE = False


class CorreosPDFProcessor:
    """Clase mejorada para procesar PDFs específicos de Correos de Costa Rica SA con múltiples patrones."""

    def __init__(self):
        """Inicializa el procesador de PDFs con patrones mejorados y flexibles."""
        if not PDFPLUMBER_AVAILABLE:
            raise ImportError("pdfplumber no está disponible. Instale con: pip install pdfplumber")

        # Múltiples patrones de búsqueda para número de factura (más flexibles)
        self.factura_patterns = [
            r'N°\s*Factura:\s*(\d{4,8})',  # Patrón original mejorado con longitud flexible
            r'No\.\s*Factura:\s*(\d{4,8})',  # Variación con "No."
            r'Número\s*Factura:\s*(\d{4,8})',  # Variación con "Número"
            r'Núm\.\s*Factura:\s*(\d{4,8})',  # Variación con "Núm."
            r'FACTURA\s*N°?\s*:?\s*(\d{4,8})',  # Variación en mayúsculas
            r'Factura\s*No\.\s*(\d{4,8})',  # Otro formato común
            r'Factura\s*#\s*(\d{4,8})',  # Con símbolo #
            r'N°\s*(\d{4,8})',  # Solo N° seguido de números
            r'No\.\s*(\d{4,8})',  # Solo No. seguido de números
            r'Documento\s*N°\s*(\d{4,8})',  # Variación con "Documento"
            r'(?:N°|No\.|Núm\.|#)\s*(\d{4,8})',  # Patrón genérico flexible
        ]

        # Patrones para códigos de guías (mantenidos del original)
        self.guias_pattern = r'Guías?\s*\n(.*?)(?:\n\n|\Z)'
        self.guia_code_pattern = r'([A-Z]{2}\d{9}[A-Z]{0,2})'

        # Configuración de debugging
        self.debug_mode = True
        self.extraction_stats = {
            'pages_processed': 0,
            'pages_failed': 0,
            'patterns_tried': 0,
            'successful_pattern': None
        }

    def find_associated_pdf(self, xml_file_path: str) -> Optional[str]:
        """
        Busca el PDF asociado al archivo XML en la misma carpeta con búsqueda mejorada.

        Args:
            xml_file_path (str): Ruta del archivo XML

        Returns:
            str: Ruta del PDF encontrado o None si no existe
        """
        try:
            # Obtener directorio del XML
            xml_dir = os.path.dirname(xml_file_path)
            xml_name = os.path.splitext(os.path.basename(xml_file_path))[0]

            self._debug_log(f"Buscando PDF para XML: {xml_name}")

            # Patrones de búsqueda de PDF más exhaustivos
            search_patterns = [
                f"{xml_name}.pdf",
                f"{xml_name}.PDF",
                f"{xml_name.lower()}.pdf",
                f"{xml_name.upper()}.pdf",
            ]

            # Buscar con nombres exactos
            for pattern in search_patterns:
                pdf_path = os.path.join(xml_dir, pattern)
                if os.path.exists(pdf_path):
                    self._debug_log(f"PDF encontrado con patrón exacto: {pattern}")
                    return pdf_path

            # Buscar cualquier PDF que contenga parte del nombre del XML
            xml_base = xml_name.split('-')[0] if '-' in xml_name else xml_name[:10]

            for file in os.listdir(xml_dir):
                if file.lower().endswith('.pdf'):
                    # Verificar si contiene parte del nombre base
                    if xml_base.lower() in file.lower():
                        pdf_path = os.path.join(xml_dir, file)
                        self._debug_log(f"PDF encontrado por coincidencia parcial: {file}")
                        return pdf_path

            # Como último recurso, buscar cualquier PDF en la carpeta
            for file in os.listdir(xml_dir):
                if file.lower().endswith('.pdf'):
                    pdf_path = os.path.join(xml_dir, file)
                    self._debug_log(f"PDF encontrado (cualquier PDF): {file}")
                    return pdf_path

            self._debug_log("No se encontró ningún PDF asociado")
            return None

        except Exception as e:
            self._debug_log(f"Error buscando PDF asociado: {e}")
            return None

    def extract_factura_number(self, pdf_text: str) -> Optional[str]:
        """
        Extrae el número de factura usando múltiples patrones mejorados.

        Args:
            pdf_text (str): Texto extraído del PDF

        Returns:
            str: Número de factura o None si no se encuentra
        """
        try:
            self._debug_log("Iniciando extracción de número de factura con patrones mejorados")

            # Limpiar el texto para mejorar la búsqueda
            cleaned_text = self._clean_text_for_search(pdf_text)

            # Probar cada patrón hasta encontrar un match
            for i, pattern in enumerate(self.factura_patterns, 1):
                self.extraction_stats['patterns_tried'] += 1

                matches = re.findall(pattern, cleaned_text, re.IGNORECASE | re.MULTILINE)

                if matches:
                    # Tomar el primer match válido
                    factura_number = matches[0]

                    # Validar que el número tenga sentido
                    if self._validate_factura_number(factura_number):
                        self.extraction_stats['successful_pattern'] = f"Patrón {i}: {pattern}"
                        self._debug_log(f"✅ Número de factura encontrado con patrón {i}: {factura_number}")
                        self._debug_log(f"Patrón exitoso: {pattern}")
                        return factura_number
                    else:
                        self._debug_log(f"⚠️ Número inválido encontrado con patrón {i}: {factura_number}")

            # Si no encontramos nada con patrones específicos, buscar números que parezcan facturas
            fallback_number = self._fallback_factura_search(cleaned_text)
            if fallback_number:
                self.extraction_stats['successful_pattern'] = "Búsqueda de respaldo"
                self._debug_log(f"✅ Número de factura encontrado con búsqueda de respaldo: {fallback_number}")
                return fallback_number

            self._debug_log("❌ No se pudo extraer número de factura con ningún patrón")
            return None

        except Exception as e:
            self._debug_log(f"❌ Error extrayendo número de factura: {e}")
            return None

    def _clean_text_for_search(self, text: str) -> str:
        """
        Limpia y normaliza el texto para mejorar la búsqueda de patrones.

        Args:
            text (str): Texto original

        Returns:
            str: Texto limpio y normalizado
        """
        try:
            # Reemplazar caracteres problemáticos
            cleaned = text.replace('\u00a0', ' ')  # Espacios no separables
            cleaned = re.sub(r'\s+', ' ', cleaned)  # Múltiples espacios a uno solo
            cleaned = cleaned.replace('º', '°')  # Normalizar símbolos de grado
            cleaned = cleaned.replace('Nº', 'N°')  # Normalizar N°
            cleaned = cleaned.replace('n°', 'N°')  # Normalizar minúsculas

            return cleaned.strip()
        except Exception as e:
            self._debug_log(f"Error limpiando texto: {e}")
            return text

    def _validate_factura_number(self, number: str) -> bool:
        """
        Valida que un número de factura tenga el formato esperado.

        Args:
            number (str): Número a validar

        Returns:
            bool: True si el número es válido
        """
        try:
            # Convertir a entero para validar
            num_int = int(number)

            # Validar longitud (entre 4 y 8 dígitos es razonable)
            if len(number) < 4 or len(number) > 8:
                return False

            # Validar que no sea cero o muy pequeño
            if num_int < 1000:
                return False

            # Validar que no sea excesivamente grande
            if num_int > 99999999:
                return False

            return True

        except ValueError:
            return False

    def _fallback_factura_search(self, text: str) -> Optional[str]:
        """
        Búsqueda de respaldo para números que podrían ser facturas.

        Args:
            text (str): Texto donde buscar

        Returns:
            str: Número de factura encontrado o None
        """
        try:
            # Buscar líneas que contengan "factura" y números
            lines = text.split('\n')

            for line in lines:
                if 'factura' in line.lower():
                    # Buscar números de 4-8 dígitos en esta línea
                    numbers = re.findall(r'\b(\d{4,8})\b', line)

                    for number in numbers:
                        if self._validate_factura_number(number):
                            self._debug_log(
                                f"Número encontrado en búsqueda de respaldo: {number} (línea: {line.strip()[:50]})")
                            return number

            # Si no encontramos nada específico, buscar números de longitud apropiada cerca de texto relevante
            factura_context_pattern = r'(?:factura|invoice|doc|documento).{0,50}?(\d{4,8})'
            matches = re.findall(factura_context_pattern, text, re.IGNORECASE | re.DOTALL)

            for match in matches:
                if self._validate_factura_number(match):
                    self._debug_log(f"Número encontrado por contexto: {match}")
                    return match

            return None

        except Exception as e:
            self._debug_log(f"Error en búsqueda de respaldo: {e}")
            return None

    def extract_guias_codes(self, pdf_text: str) -> List[str]:
        """
        Extrae los códigos de guías con validación mejorada.

        Args:
            pdf_text (str): Texto extraído del PDF

        Returns:
            List[str]: Lista de códigos de guías encontrados
        """
        try:
            self._debug_log("Iniciando extracción de códigos de guías")
            guias_codes = []

            # Buscar todas las ocurrencias del patrón de código de guía en todo el texto
            matches = re.findall(self.guia_code_pattern, pdf_text)

            for match in matches:
                # Validar que el código tenga el formato esperado
                if self._validate_guia_code(match):
                    guias_codes.append(match)

            # Eliminar duplicados manteniendo el orden
            seen = set()
            unique_guias = []
            for guia in guias_codes:
                if guia not in seen:
                    seen.add(guia)
                    unique_guias.append(guia)

            self._debug_log(f"Códigos de guías encontrados: {len(unique_guias)} únicos de {len(matches)} totales")
            for i, code in enumerate(unique_guias, 1):
                self._debug_log(f"  {i}. {code}")

            return unique_guias

        except Exception as e:
            self._debug_log(f"Error extrayendo códigos de guías: {e}")
            return []

    def _validate_guia_code(self, code: str) -> bool:
        """
        Valida que un código de guía tenga el formato esperado con validación mejorada.

        Args:
            code (str): Código a validar

        Returns:
            bool: True si el código es válido
        """
        try:
            # Formato esperado: 2 letras + 9 dígitos + 2 letras opcionales
            # Ejemplos: NE084204615CR, NE116467408CR, NE123456789
            pattern = r'^[A-Z]{2}\d{9}[A-Z]{0,2}$'
            is_valid = bool(re.match(pattern, code))

            if not is_valid:
                self._debug_log(f"Código de guía inválido: {code}")

            return is_valid

        except Exception:
            return False

    def extract_pdf_text(self, pdf_path: str) -> Optional[str]:
        """
        Extrae todo el texto de un archivo PDF con manejo robusto de errores.

        Args:
            pdf_path (str): Ruta del archivo PDF

        Returns:
            str: Texto extraído del PDF o None si hay error
        """
        try:
            if not os.path.exists(pdf_path):
                self._debug_log(f"❌ PDF no encontrado: {pdf_path}")
                return None

            self._debug_log(f"📄 Iniciando extracción de texto de PDF: {os.path.basename(pdf_path)}")
            full_text = ""
            self.extraction_stats['pages_processed'] = 0
            self.extraction_stats['pages_failed'] = 0

            with pdfplumber.open(pdf_path) as pdf:
                total_pages = len(pdf.pages)
                self._debug_log(f"Total de páginas en PDF: {total_pages}")

                for page_num, page in enumerate(pdf.pages, 1):
                    try:
                        self.extraction_stats['pages_processed'] += 1
                        page_text = page.extract_text()

                        if page_text:
                            full_text += page_text + "\n"
                            self._debug_log(f"✅ Página {page_num}: {len(page_text)} caracteres extraídos")
                        else:
                            self._debug_log(f"⚠️ Página {page_num}: Sin texto extraído")
                            self.extraction_stats['pages_failed'] += 1

                    except Exception as e:
                        self.extraction_stats['pages_failed'] += 1
                        self._debug_log(f"❌ Error extrayendo texto de página {page_num}: {e}")

                        # Intentar métodos alternativos para esta página
                        try:
                            # Método alternativo: extraer texto caracter por caracter
                            chars = page.chars
                            if chars:
                                alt_text = ''.join([char.get('text', '') for char in chars])
                                if alt_text.strip():
                                    full_text += alt_text + "\n"
                                    self._debug_log(f"✅ Página {page_num}: Extraído con método alternativo")
                                    continue
                        except Exception as alt_e:
                            self._debug_log(f"❌ Método alternativo también falló para página {page_num}: {alt_e}")

                        continue

            # Estadísticas finales
            success_rate = ((self.extraction_stats['pages_processed'] - self.extraction_stats['pages_failed']) /
                            self.extraction_stats['pages_processed'] * 100) if self.extraction_stats[
                                                                                   'pages_processed'] > 0 else 0

            self._debug_log(f"📊 Extracción completada:")
            self._debug_log(f"  - Páginas procesadas: {self.extraction_stats['pages_processed']}")
            self._debug_log(f"  - Páginas fallidas: {self.extraction_stats['pages_failed']}")
            self._debug_log(f"  - Tasa de éxito: {success_rate:.1f}%")
            self._debug_log(f"  - Texto total extraído: {len(full_text)} caracteres")

            if not full_text.strip():
                self._debug_log("❌ No se pudo extraer texto del PDF")
                return None

            return full_text.strip()

        except Exception as e:
            self._debug_log(f"❌ Error crítico abriendo PDF {pdf_path}: {e}")
            return None

    def process_correos_pdf(self, xml_file_path: str) -> Tuple[bool, Optional[str]]:
        """
        Procesa un PDF de Correos de Costa Rica SA con manejo mejorado y extrae la información necesaria.

        Args:
            xml_file_path (str): Ruta del archivo XML asociado

        Returns:
            Tuple[bool, str]: (éxito, dato_formateado) donde dato_formateado es como:
                             "(345520) SERVICIO EMS (ENVIO DE PAQUETES)/GUIA NE084204615CR"
        """
        try:
            self._debug_log(f"🚀 Iniciando procesamiento de PDF de Correos para XML: {os.path.basename(xml_file_path)}")

            # Resetear estadísticas
            self.extraction_stats = {
                'pages_processed': 0,
                'pages_failed': 0,
                'patterns_tried': 0,
                'successful_pattern': None
            }

            # Buscar PDF asociado
            pdf_path = self.find_associated_pdf(xml_file_path)
            if not pdf_path:
                self._debug_log("❌ No se encontró PDF asociado")
                return False, None

            self._debug_log(f"📎 PDF asociado encontrado: {os.path.basename(pdf_path)}")

            # Extraer texto del PDF
            pdf_text = self.extract_pdf_text(pdf_path)
            if not pdf_text:
                self._debug_log("❌ No se pudo extraer texto del PDF")
                return False, None

            # Mostrar muestra del texto extraído para debugging
            self._debug_log(f"📝 Muestra del texto extraído (primeros 200 caracteres):")
            self._debug_log(f"  {pdf_text[:200]}...")

            # Extraer número de factura
            factura_number = self.extract_factura_number(pdf_text)
            if not factura_number:
                self._debug_log("❌ No se encontró número de factura")
                return False, None

            # Extraer códigos de guías
            guias_codes = self.extract_guias_codes(pdf_text)
            if not guias_codes:
                self._debug_log("❌ No se encontraron códigos de guías")
                return False, None

            # Formatear el resultado
            guias_text = " - ".join(guias_codes)
            formatted_data = f"({factura_number}) SERVICIO EMS (ENVIO DE PAQUETES)/GUIA {guias_text}"

            # Log de éxito con estadísticas
            self._debug_log("🎯 PROCESAMIENTO EXITOSO:")
            self._debug_log(f"  - Número de factura: {factura_number}")
            self._debug_log(f"  - Códigos de guías: {len(guias_codes)} encontrados")
            self._debug_log(f"  - Patrón exitoso: {self.extraction_stats['successful_pattern']}")
            self._debug_log(f"  - Resultado final: {formatted_data}")

            return True, formatted_data

        except Exception as e:
            self._debug_log(f"❌ Error crítico procesando PDF de Correos: {e}")
            return False, None

    def _debug_log(self, message: str):
        """
        Método de logging para debugging con control de verbosidad.

        Args:
            message (str): Mensaje a loggear
        """
        if self.debug_mode:
            print(f"[PDF_PROCESSOR] {message}")

    def set_debug_mode(self, enabled: bool):
        """
        Habilita o deshabilita el modo debug.

        Args:
            enabled (bool): True para habilitar debug
        """
        self.debug_mode = enabled

    def get_extraction_stats(self) -> Dict:
        """
        Obtiene las estadísticas de la última extracción.

        Returns:
            Dict: Estadísticas de extracción
        """
        return self.extraction_stats.copy()

    def test_pdf_extraction(self, pdf_path: str) -> None:
        """
        Método de prueba mejorado para verificar la extracción de datos de un PDF.

        Args:
            pdf_path (str): Ruta del PDF a probar
        """
        try:
            print(f"=== PRUEBA AVANZADA DE EXTRACCIÓN PDF ===")
            print(f"Archivo: {pdf_path}")
            print()

            if not os.path.exists(pdf_path):
                print("❌ El archivo PDF no existe")
                return

            # Habilitar modo debug temporalmente
            original_debug = self.debug_mode
            self.debug_mode = True

            # Extraer texto
            print("📄 Extrayendo texto del PDF...")
            pdf_text = self.extract_pdf_text(pdf_path)

            if not pdf_text:
                print("❌ No se pudo extraer texto del PDF")
                return

            print("✅ Texto extraído exitosamente")
            print(f"Longitud del texto: {len(pdf_text)} caracteres")
            print()

            # Mostrar primeras líneas del texto
            lines = pdf_text.split('\n')[:15]
            print("📝 Primeras 15 líneas del texto:")
            for i, line in enumerate(lines, 1):
                print(f"  {i:2d}. {line}")
            print()

            # Buscar número de factura con todos los patrones
            print("🔍 Probando todos los patrones de número de factura...")
            factura_number = self.extract_factura_number(pdf_text)

            if factura_number:
                print(f"✅ Número de factura encontrado: {factura_number}")
                print(f"✅ Patrón exitoso: {self.extraction_stats['successful_pattern']}")
            else:
                print("❌ No se encontró número de factura")
                print(f"Patrones probados: {self.extraction_stats['patterns_tried']}")
            print()

            # Buscar códigos de guías
            print("🔍 Buscando códigos de guías...")
            guias_codes = self.extract_guias_codes(pdf_text)
            if guias_codes:
                print(f"✅ Códigos de guías encontrados: {len(guias_codes)}")
                for i, code in enumerate(guias_codes, 1):
                    print(f"  {i}. {code}")
            else:
                print("❌ No se encontraron códigos de guías")
            print()

            # Resultado final
            if factura_number and guias_codes:
                guias_text = " - ".join(guias_codes)
                final_result = f"({factura_number}) SERVICIO EMS (ENVIO DE PAQUETES)/GUIA {guias_text}"
                print("🎯 RESULTADO FINAL:")
                print(f"   {final_result}")
            else:
                print("❌ No se pudo generar resultado final")

            # Mostrar estadísticas
            stats = self.get_extraction_stats()
            print(f"\n📊 ESTADÍSTICAS:")
            print(f"  - Páginas procesadas: {stats['pages_processed']}")
            print(f"  - Páginas fallidas: {stats['pages_failed']}")
            print(f"  - Patrones probados: {stats['patterns_tried']}")

            print("\n=== FIN DE PRUEBA ===")

            # Restaurar modo debug original
            self.debug_mode = original_debug

        except Exception as e:
            print(f"❌ Error en prueba: {e}")


def create_correos_pdf_processor() -> Optional[CorreosPDFProcessor]:
    """
    Crea una instancia del procesador mejorado de PDFs de Correos si está disponible.

    Returns:
        CorreosPDFProcessor: Instancia del procesador o None si no está disponible
    """
    try:
        return CorreosPDFProcessor()
    except ImportError:
        print("Advertencia: pdfplumber no está disponible.")
        print("Para habilitar procesamiento de PDFs, instale: pip install pdfplumber")
        return None


if __name__ == "__main__":
    # Prueba básica del procesador mejorado
    processor = create_correos_pdf_processor()

    if processor:
        print("✅ CorreosPDFProcessor mejorado inicializado correctamente")
        print("🔧 Mejoras incluidas:")
        print("  - Múltiples patrones de búsqueda para número de factura")
        print("  - Manejo robusto de errores de extracción de páginas")
        print("  - Validación mejorada de números de factura (4-8 dígitos)")
        print("  - Búsqueda de respaldo para casos difíciles")
        print("  - Logs detallados para debugging")
        print("  - Estadísticas de extracción")

        # Aquí puedes agregar una ruta de PDF para probar
        # processor.test_pdf_extraction("ruta/al/archivo.pdf")
    else:
        print("❌ No se pudo inicializar CorreosPDFProcessor")