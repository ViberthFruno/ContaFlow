# otro_texto_processor.py
"""
Procesador especializado para extraer códigos de placas desde el campo <OtroTexto> de archivos XML.
Maneja múltiples patrones de placas vehiculares y formatea la salida para ContaFlow.
"""

import re
from typing import Optional, List, Tuple


class OtroTextoProcessor:
    """Clase especializada para procesar el campo <OtroTexto> y extraer códigos de placas vehiculares."""

    def __init__(self):
        """Inicializa el procesador con los patrones de búsqueda de placas."""
        # Patrones de placas a buscar - ORDEN CORREGIDO: M/m primero, luego 6 dígitos
        self.placa_patterns = [
            # Patrón 1: 7 dígitos empezando con M/m (PRIMERO para evitar que se pierda la M)
            # Ejemplos: m914559, M 782308, M914559
            r'([mM][\s]?\d{6})',

            # Patrón 2: Formato especial CL + 6 dígitos
            # Ejemplos: CL435475
            r'(CL\d{6})',

            # Patrón 3: 6 dígitos alfanuméricos (con espacios o guiones opcionales)
            # Ejemplos: BJX 894, 123456, BJM-653, ABC123
            # IMPORTANTE: Excluir KM explícitamente
            r'((?![kK][mM][\s\-])[A-Z]{2,3}[\s\-]?\d{3,4}|\d{6}|[A-Z]{3}\d{3})'
        ]

        # Palabras clave que indican que viene una placa (AGREGADO pl:)
        self.placa_keywords = [
            r'placa\s*:',
            r'PLACA\s*:',
            r'Placa\s*:',
            r'placa\s*=',
            r'PLACA\s*=',
            r'Placa\s*=',
            r'pl\s*:',  # NUEVO: para casos como "pl:m833753"
            r'PL\s*:'  # NUEVO: versión mayúscula
        ]

        # Patrón para detectar kilometraje
        self.km_pattern = re.compile(r'[kK][mM][\s:]?\d+', re.IGNORECASE)

        # Compilar patrones para mejor rendimiento
        self.compiled_patterns = [re.compile(pattern, re.IGNORECASE) for pattern in self.placa_patterns]
        self.compiled_keywords = [re.compile(keyword, re.IGNORECASE) for keyword in self.placa_keywords]

    def extract_placa_code(self, otro_texto: str) -> Optional[str]:
        """
        Extrae el código de placa desde el campo OtroTexto.

        Args:
            otro_texto (str): Contenido del campo <OtroTexto>

        Returns:
            str: Código de placa extraído o None si no se encuentra
        """
        if not otro_texto or not otro_texto.strip():
            return None

        try:
            # Limpiar el texto
            texto_limpio = otro_texto.strip()

            # Buscar códigos de placa después de palabras clave
            placa_code = self._find_placa_after_keywords(texto_limpio)

            if placa_code:
                # Verificar que no sea un código KM
                if not self._is_km_code(placa_code):
                    return self._clean_placa_code(placa_code)

            # Si no encuentra después de keywords, buscar patrones en todo el texto
            placa_code = self._find_placa_patterns(texto_limpio)

            if placa_code:
                # Verificar que no sea un código KM
                if not self._is_km_code(placa_code):
                    return self._clean_placa_code(placa_code)

            return None

        except Exception as e:
            print(f"Error extrayendo código de placa: {e}")
            return None

    def _is_km_code(self, code: str) -> bool:
        """
        Verifica si un código extraído es en realidad un código de kilometraje.

        Args:
            code (str): Código a verificar

        Returns:
            bool: True si es un código KM
        """
        if not code:
            return False

        # Verificar si el código coincide con patrón KM
        return bool(re.match(r'^[kK][mM][\s]?\d+$', code.strip()))

    def _find_placa_after_keywords(self, texto: str) -> Optional[str]:
        """
        Busca códigos de placa después de palabras clave específicas.

        Args:
            texto (str): Texto donde buscar

        Returns:
            str: Código encontrado o None
        """
        for keyword_pattern in self.compiled_keywords:
            # Buscar la palabra clave
            keyword_match = keyword_pattern.search(texto)

            if keyword_match:
                # Obtener el texto después de la palabra clave
                start_pos = keyword_match.end()
                texto_despues = texto[start_pos:start_pos + 50]  # Buscar en los siguientes 50 caracteres

                # Buscar patrones de placa en el texto después de la palabra clave
                for pattern in self.compiled_patterns:
                    match = pattern.search(texto_despues)
                    if match:
                        return match.group(1)

        return None

    def _find_placa_patterns(self, texto: str) -> Optional[str]:
        """
        Busca patrones de placa en todo el texto.

        Args:
            texto (str): Texto donde buscar

        Returns:
            str: Código encontrado o None
        """
        for pattern in self.compiled_patterns:
            match = pattern.search(texto)
            if match:
                return match.group(1)
        return None

    def _is_only_km_info(self, otro_texto: str) -> bool:
        """
        Verifica si el OtroTexto SOLO contiene información de kilometraje.

        Args:
            otro_texto (str): Contenido del campo <OtroTexto>

        Returns:
            bool: True si solo contiene información de KM
        """
        if not otro_texto or not otro_texto.strip():
            return False

        try:
            texto_limpio = otro_texto.strip()

            # Verificar si contiene patrón KM
            km_match = self.km_pattern.search(texto_limpio)
            if not km_match:
                return False

            # Verificar que NO contenga indicadores de placa
            for keyword_pattern in self.compiled_keywords:
                if keyword_pattern.search(texto_limpio):
                    return False  # Si tiene keywords de placa, no es "solo KM"

            # Remover TODOS los patrones KM del texto
            texto_sin_km = self.km_pattern.sub('', texto_limpio)

            # Limpiar espacios, puntuación básica y caracteres comunes
            texto_sin_km = re.sub(r'[:\s\-_.,;]+', '', texto_sin_km).strip()

            # Si después de remover KM y limpiar queda muy poco texto significativo
            # consideramos que es "solo KM"
            return len(texto_sin_km) < 5

        except Exception:
            return False

    def _clean_placa_code(self, placa_code: str) -> str:
        """
        Limpia y normaliza el código de placa extraído.

        Args:
            placa_code (str): Código bruto extraído

        Returns:
            str: Código limpio y normalizado
        """
        if not placa_code:
            return ""

        # Limpiar espacios extra y caracteres especiales
        cleaned = placa_code.strip()

        # Normalizar espacios múltiples a uno solo
        cleaned = re.sub(r'\s+', ' ', cleaned)

        # Remover caracteres no alfanuméricos excepto espacios y guiones
        cleaned = re.sub(r'[^\w\s\-]', '', cleaned)

        return cleaned.upper()  # Convertir a mayúsculas para consistencia

    def format_combustible_output(self, placa_code: str) -> str:
        """
        Formatea el código de placa en el formato requerido para ContaFlow.

        Args:
            placa_code (str): Código de placa extraído

        Returns:
            str: Texto formateado como "Combustible / Placa: [CÓDIGO]"
        """
        if not placa_code or not placa_code.strip():
            return ""

        return f"Combustible / Placa: {placa_code.strip()}"

    def process_otro_texto(self, otro_texto: str) -> Optional[str]:
        """
        Proceso principal que extrae y formatea el código de placa desde OtroTexto.

        Args:
            otro_texto (str): Contenido del campo <OtroTexto>

        Returns:
            str: Texto formateado listo para usar o None si no se encuentra placa
        """
        # Primero intentar extraer código de placa normal
        placa_code = self.extract_placa_code(otro_texto)

        if placa_code:
            return self.format_combustible_output(placa_code)

        # Si no encontró placa, verificar si es caso especial de "solo KM"
        if self._is_only_km_info(otro_texto):
            return "Combustible / Placa: ?"

        # Si no es placa ni caso especial KM, retornar None para usar Detalle
        return None

    def get_extraction_stats(self, otro_texto_list: List[str]) -> dict:
        """
        Obtiene estadísticas de extracción para una lista de textos OtroTexto.

        Args:
            otro_texto_list (List[str]): Lista de textos a analizar

        Returns:
            dict: Estadísticas de extracción
        """
        stats = {
            'total_processed': len(otro_texto_list),
            'placas_found': 0,
            'km_only_cases': 0,
            'placas_not_found': 0,
            'extraction_rate': 0.0,
            'patterns_found': {
                'alphanumeric_6': 0,
                'M_format_7': 0,
                'CL_special': 0,
                'km_only': 0,
                'other': 0
            }
        }

        for texto in otro_texto_list:
            result = self.process_otro_texto(texto)

            if result:
                if result == "Combustible / Placa: ?":
                    stats['km_only_cases'] += 1
                    stats['patterns_found']['km_only'] += 1
                else:
                    stats['placas_found'] += 1
                    placa_code = result.replace("Combustible / Placa: ", "")

                    # Clasificar el tipo de patrón encontrado
                    if re.match(r'[mM][\s]?\d{6}', placa_code, re.IGNORECASE):
                        stats['patterns_found']['M_format_7'] += 1
                    elif re.match(r'CL\d{6}', placa_code, re.IGNORECASE):
                        stats['patterns_found']['CL_special'] += 1
                    elif re.match(r'([A-Z]{2,3}[\s\-]?\d{3,4}|\d{6}|[A-Z]{3}\d{3})', placa_code, re.IGNORECASE):
                        stats['patterns_found']['alphanumeric_6'] += 1
                    else:
                        stats['patterns_found']['other'] += 1
            else:
                stats['placas_not_found'] += 1

        # Calcular tasa de extracción (incluyendo casos KM)
        successful_extractions = stats['placas_found'] + stats['km_only_cases']
        if stats['total_processed'] > 0:
            stats['extraction_rate'] = (successful_extractions / stats['total_processed']) * 100

        return stats

    def validate_placa_format(self, placa_code: str) -> Tuple[bool, str]:
        """
        Valida si un código de placa cumple con los formatos esperados.

        Args:
            placa_code (str): Código a validar

        Returns:
            Tuple[bool, str]: (es_válido, descripción_del_formato)
        """
        try:
            if not placa_code or not placa_code.strip():
                return False, "Código vacío"

            cleaned_code = placa_code.strip().upper()

            # Caso especial para KM
            if cleaned_code == "?":
                return True, "Caso especial KM"

            # Verificar que no sea código KM
            if self._is_km_code(cleaned_code):
                return False, "Es código de kilometraje, no placa"

            # Validar patrón 1: 6 dígitos alfanuméricos
            if re.match(r'^([A-Z]{2,3}[\s\-]?\d{3,4}|\d{6}|[A-Z]{3}\d{3})$', cleaned_code):
                return True, "6 dígitos alfanuméricos"

            # Validar patrón 2: 7 dígitos con M
            if re.match(r'^[mM][\s]?\d{6}$', cleaned_code):
                return True, "7 dígitos formato M"

            # Validar patrón 3: Formato CL especial
            if re.match(r'^CL\d{6}$', cleaned_code):
                return True, "Formato especial CL"

            return False, "Formato no reconocido"
        except Exception:
            return False, "Error validando formato"


def create_otro_texto_processor():
    """
    Factory function para crear una instancia del procesador de OtroTexto.

    Returns:
        OtroTextoProcessor: Instancia del procesador
    """
    return OtroTextoProcessor()


# Funciones de utilidad para uso directo
def extract_placa_from_otro_texto(otro_texto: str) -> Optional[str]:
    """
    Función de utilidad para extraer placa directamente desde OtroTexto.

    Args:
        otro_texto (str): Contenido del campo <OtroTexto>

    Returns:
        str: Texto formateado "Combustible / Placa: [CÓDIGO]" o None
    """
    processor = create_otro_texto_processor()
    return processor.process_otro_texto(otro_texto)


def validate_placa_code(placa_code: str) -> bool:
    """
    Función de utilidad para validar un código de placa.

    Args:
        placa_code (str): Código a validar

    Returns:
        bool: True si el formato es válido
    """
    processor = create_otro_texto_processor()
    is_valid, _ = processor.validate_placa_format(placa_code)
    return is_valid


# Ejemplo de uso si se ejecuta directamente
if __name__ == "__main__":
    # Crear procesador
    processor = create_otro_texto_processor()

    # Ejemplos de prueba corregidos
    test_cases = [
        """km:40800 pl:m833753""",  # Debería extraer m833753

        """KM 9962""",  # Debería ser caso especial: ?

        """Placa:BJX 894 KM 9509""",  # Debería extraer BJX 894

        """Factura Contado:706916 ID Despacho:775572 Fecha:1/7/2025 10:36:02 a. m. Posicion:7 Pistero: JOSUE SOTO CARRILLO Placa:m914559 Kilometraje:20,169 KM/L:380.007 Orden Compra: Comentario:""",
        # Debería extraer m914559

        """KM: 8765""",  # Caso especial: solo kilometraje

        """Ejemplo sin placa válida o sin campo placa"""
    ]

    print("🧪 PRUEBAS DEL PROCESADOR DE OTRO TEXTO - KM CORREGIDO")
    print("=" * 60)

    for i, test_case in enumerate(test_cases, 1):
        print(f"\n📝 Caso de prueba {i}:")
        print(f"Texto: {test_case}")

        result = processor.process_otro_texto(test_case)

        if result:
            if result == "Combustible / Placa: ?":
                print(f"🔧 Resultado (caso KM): {result}")
            else:
                print(f"✅ Resultado: {result}")
        else:
            print("❌ No se encontró código de placa (usará Detalle)")

    # Estadísticas
    print(f"\n📊 ESTADÍSTICAS:")
    stats = processor.get_extraction_stats(test_cases)
    print(f"Total procesados: {stats['total_processed']}")
    print(f"Placas encontradas: {stats['placas_found']}")
    print(f"Casos solo KM: {stats['km_only_cases']}")
    print(f"Sin extracción: {stats['placas_not_found']}")
    print(f"Tasa de extracción: {stats['extraction_rate']:.1f}%")