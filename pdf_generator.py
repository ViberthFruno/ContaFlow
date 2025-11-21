# pdf_generator.py
"""
Generador de PDF profesional para ContaFlow con reporte de XMLs excluidos por fecha y subsección de Correos.
Crea informes PDF limpios y modernos con resúmenes del procesamiento de archivos empresariales,
detalles de XMLs que no se procesaron por estar fuera del mes actual y una subsección especial
para archivos de Correos de Costa Rica SA procesados.
"""

import os
from datetime import datetime
from typing import Dict, List, Optional
from collections import defaultdict

try:
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import letter, A4
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.units import inch
    from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT
    from reportlab.platypus.flowables import HRFlowable

    REPORTLAB_AVAILABLE = True
except ImportError:
    REPORTLAB_AVAILABLE = False


class PDFGenerator:
    """Clase para generar PDF profesionales y limpios del procesamiento con información de filtrado por fecha y Correos."""

    def __init__(self):
        """Inicializa el generador de PDF con diseño limpio y consistente."""
        if not REPORTLAB_AVAILABLE:
            raise ImportError("reportlab no está disponible. Instale con: pip install reportlab")

        self.page_width, self.page_height = A4
        self.margin = 0.8 * inch

        # Paleta de colores consistente (solo azules)
        self.colors = {
            'primary_blue': colors.HexColor('#0D47A1'),  # Azul principal para todo
            'secondary_blue': colors.HexColor('#1565C0'),  # Azul para headers de tabla
            'light_blue': colors.HexColor('#E8EAF6'),  # Azul muy claro para filas
            'white': colors.white,
            'text_dark': colors.HexColor('#263238'),  # Texto oscuro
            'text_light': colors.HexColor('#546E7A'),  # Texto secundario
            'border_light': colors.HexColor('#E0E0E0'),  # Bordes discretos
            'warning_orange': colors.HexColor('#FF9800'),  # Para XMLs excluidos
            'light_orange': colors.HexColor('#FFF3E0'),  # Fondo para XMLs excluidos
            'correos_green': colors.HexColor('#4CAF50'),  # Verde para sección Correos
            'light_green': colors.HexColor('#E8F5E8')  # Verde claro para Correos
        }

        # Configurar estilos
        self.styles = getSampleStyleSheet()
        self._setup_clean_styles()

    def _setup_clean_styles(self):
        """Configura estilos limpios y consistentes para el PDF."""

        # Logo Fruno principal
        self.styles.add(ParagraphStyle(
            name='FrunoLogo',
            parent=self.styles['Normal'],
            fontSize=32,
            textColor=self.colors['primary_blue'],
            alignment=TA_CENTER,
            fontName='Helvetica-Bold',
            spaceAfter=20,
            spaceBefore=10
        ))

        # Título del reporte
        self.styles.add(ParagraphStyle(
            name='ReportTitle',
            parent=self.styles['Title'],
            fontSize=18,
            textColor=self.colors['primary_blue'],
            spaceAfter=8,
            alignment=TA_CENTER,
            fontName='Helvetica-Bold'
        ))

        # Subtítulo del reporte
        self.styles.add(ParagraphStyle(
            name='ReportSubtitle',
            parent=self.styles['Normal'],
            fontSize=10,
            textColor=self.colors['text_light'],
            alignment=TA_CENTER,
            fontName='Helvetica',
            spaceAfter=30
        ))

        # Título de sección principal
        self.styles.add(ParagraphStyle(
            name='MainSectionTitle',
            parent=self.styles['Heading1'],
            fontSize=16,
            textColor=self.colors['primary_blue'],
            spaceBefore=20,
            spaceAfter=15,
            fontName='Helvetica-Bold',
            alignment=TA_CENTER
        ))

        # Título de sección secundaria (Para XMLs excluidos y Correos)
        self.styles.add(ParagraphStyle(
            name='SecondarySectionTitle',
            parent=self.styles['Heading2'],
            fontSize=14,
            textColor=self.colors['warning_orange'],
            spaceBefore=15,
            spaceAfter=12,
            fontName='Helvetica-Bold',
            alignment=TA_CENTER
        ))

        # Título especial para sección de Correos
        self.styles.add(ParagraphStyle(
            name='CorreosSectionTitle',
            parent=self.styles['Heading2'],
            fontSize=14,
            textColor=self.colors['correos_green'],
            spaceBefore=15,
            spaceAfter=12,
            fontName='Helvetica-Bold',
            alignment=TA_CENTER
        ))

        # Título de empresa para XMLs excluidos
        self.styles.add(ParagraphStyle(
            name='CompanyTitle',
            parent=self.styles['Normal'],
            fontSize=12,
            textColor=self.colors['secondary_blue'],
            spaceBefore=15,
            spaceAfter=10,
            fontName='Helvetica-Bold',
            alignment=TA_LEFT
        ))

        # Footer limpio
        self.styles.add(ParagraphStyle(
            name='CleanFooter',
            parent=self.styles['Normal'],
            fontSize=8,
            textColor=self.colors['text_light'],
            alignment=TA_CENTER,
            fontName='Helvetica'
        ))

        # Texto informativo para secciones
        self.styles.add(ParagraphStyle(
            name='InfoText',
            parent=self.styles['Normal'],
            fontSize=10,
            textColor=self.colors['text_light'],
            alignment=TA_CENTER,
            fontName='Helvetica',
            spaceAfter=15
        ))

    def generate_summary_pdf(self, data: Dict, output_path: str) -> bool:
        """
        Genera un PDF limpio y profesional con información de procesamiento, XMLs excluidos y subsección de Correos.

        Args:
            data (Dict): Datos del procesamiento incluyendo XMLs excluidos y estadísticas de Correos
            output_path (str): Ruta donde guardar el PDF

        Returns:
            bool: True si se generó exitosamente
        """
        try:
            # Crear directorio si no existe
            os.makedirs(os.path.dirname(output_path), exist_ok=True)

            # Crear documento PDF
            doc = SimpleDocTemplate(
                output_path,
                pagesize=A4,
                rightMargin=self.margin,
                leftMargin=self.margin,
                topMargin=self.margin * 0.7,
                bottomMargin=self.margin * 0.7
            )

            # Construir contenido del PDF
            story = []

            # Header principal
            self._add_clean_header(story, data)

            # Resumen estadístico con información de filtrado
            self._add_statistics_summary(story, data)

            # Tabla de empresas (contenido principal)
            self._add_companies_table_clean(story, data)

            # NUEVA SUBSECCIÓN: Procesamiento de Correos de Costa Rica SA
            self._add_correos_processing_section(story, data)

            # Footer de la primera página
            self._add_clean_footer(story)

            # PÁGINA NUEVA para XMLs excluidos
            story.append(PageBreak())

            # Sección de XMLs excluidos por fecha (en nueva página)
            self._add_excluded_xmls_section(story, data)

            # Construir PDF
            doc.build(story)
            return True

        except Exception as e:
            print(f"Error generando PDF: {e}")
            return False

    def _add_clean_header(self, story: List, data: Dict):
        """Agrega un header limpio y profesional."""

        # Logo Fruno prominente
        story.append(Paragraph("FRUNO", self.styles['FrunoLogo']))

        # Título del reporte
        story.append(Paragraph("Reporte de Contabilidad", self.styles['ReportTitle']))

        # Información de fecha y sistema
        timestamp = data.get('timestamp', datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        current_date = datetime.now()
        subtitle = f"Generado el {timestamp} • Procesado por ContaFlow • Mes actual: {current_date.month:02d}/{current_date.year}"
        story.append(Paragraph(subtitle, self.styles['ReportSubtitle']))

    def _add_statistics_summary(self, story: List, data: Dict):
        """Agrega un resumen estadístico con información de filtrado por fecha y Correos."""

        # Obtener estadísticas del procesamiento
        stats = data.get('processing_stats', {})

        if not stats:
            return

        # Título de estadísticas
        story.append(Paragraph("Resumen de Procesamiento", self.styles['MainSectionTitle']))

        # Crear tabla de estadísticas en dos columnas incluyendo estadísticas de Correos
        stats_data = [
            ['Métrica', 'Cantidad'],
            ['XMLs encontrados total', f"{stats.get('total_xml_count', 0):,}"],
            ['XMLs del mes actual', f"{stats.get('total_xml_current_month', 0):,}"],
            ['XMLs excluidos por fecha', f"{stats.get('total_xml_excluded_by_date', 0):,}"],
            ['Matches exitosos', f"{stats.get('total_matches', 0):,}"],
            ['Archivos Excel generados', f"{stats.get('files_created', 0):,}"],
            ['Empresas procesadas', f"{stats.get('companies_processed', 0):,}"],
            ['PDFs de Correos procesados', f"{stats.get('correos_pdfs_processed', 0):,}"],
            ['Matches de Correos generados', f"{stats.get('correos_matches', 0):,}"]
        ]

        # Crear tabla con diseño limpio
        stats_table = Table(stats_data, colWidths=[3.5 * inch, 2.0 * inch])

        stats_table.setStyle(TableStyle([
            # Header
            ('BACKGROUND', (0, 0), (-1, 0), self.colors['primary_blue']),
            ('TEXTCOLOR', (0, 0), (-1, 0), self.colors['white']),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 11),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('TOPPADDING', (0, 0), (-1, 0), 12),

            # Body
            ('BACKGROUND', (0, 1), (-1, -1), self.colors['white']),
            ('TEXTCOLOR', (0, 1), (-1, -1), self.colors['text_dark']),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 10),
            ('TOPPADDING', (0, 1), (-1, -1), 10),
            ('LEFTPADDING', (0, 0), (-1, -1), 15),
            ('RIGHTPADDING', (0, 0), (-1, -1), 15),

            # Alineación
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),
            ('ALIGN', (1, 1), (-1, -1), 'CENTER'),

            # Filas alternadas
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [self.colors['white'], self.colors['light_blue']]),

            # Bordes
            ('GRID', (0, 0), (-1, -1), 1, self.colors['border_light']),
            ('LINEBELOW', (0, 0), (-1, 0), 2, self.colors['secondary_blue']),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),

            # Resaltar filas de Correos con color verde claro
            ('ROWBACKGROUNDS', (0, 7), (-1, 8), [self.colors['light_green'], self.colors['light_green']]),
        ]))

        story.append(stats_table)
        story.append(Spacer(1, 30))

    def _add_companies_table_clean(self, story: List, data: Dict):
        """Agrega una tabla limpia y optimizada de empresas."""

        # Título de la sección
        story.append(Paragraph("Archivos Procesados por Empresa", self.styles['MainSectionTitle']))

        files = data.get('files', [])
        if not files:
            no_data = Paragraph("No hay archivos procesados para mostrar.", self.styles['InfoText'])
            story.append(no_data)
            return

        # Agrupar estadísticas por empresa
        company_stats = {}
        for file_info in files:
            company = file_info.get('company_name', 'Empresa Desconocida')
            if company not in company_stats:
                company_stats[company] = {
                    'files': 0, 'matches': 0, 'manual_reviews': 0
                }

            company_stats[company]['files'] += 1
            company_stats[company]['matches'] += file_info.get('matches', 0)
            company_stats[company]['manual_reviews'] += file_info.get('manual_reviews', 0)

        # Crear tabla con diseño limpio y consistente
        table_data = [['Empresa', 'Archivos', 'Matches', 'Revision Manual']]

        for company, stats in company_stats.items():
            # Truncar nombre de empresa si es muy largo
            display_name = company if len(company) <= 30 else company[:27] + "..."

            table_data.append([
                display_name,
                str(stats['files']),
                f"{stats['matches']:,}",
                str(stats['manual_reviews'])
            ])

        # Crear tabla con columnas optimizadas
        table = Table(table_data, colWidths=[2.6 * inch, 1.0 * inch, 1.3 * inch, 1.4 * inch])

        # Aplicar estilo limpio y consistente
        table.setStyle(TableStyle([
            # Header con color principal consistente
            ('BACKGROUND', (0, 0), (-1, 0), self.colors['primary_blue']),
            ('TEXTCOLOR', (0, 0), (-1, 0), self.colors['white']),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 11),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 15),
            ('TOPPADDING', (0, 0), (-1, 0), 15),

            # Body con colores consistentes
            ('BACKGROUND', (0, 1), (-1, -1), self.colors['white']),
            ('TEXTCOLOR', (0, 1), (-1, -1), self.colors['text_dark']),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 12),
            ('TOPPADDING', (0, 1), (-1, -1), 12),
            ('LEFTPADDING', (0, 0), (-1, -1), 15),
            ('RIGHTPADDING', (0, 0), (-1, -1), 15),

            # Alineación optimizada
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),  # Nombres de empresas a la izquierda
            ('ALIGN', (1, 1), (-1, -1), 'CENTER'),  # Números centrados

            # Filas alternadas con color consistente
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [self.colors['white'], self.colors['light_blue']]),

            # Bordes limpios y discretos
            ('GRID', (0, 0), (-1, -1), 1, self.colors['border_light']),
            ('LINEBELOW', (0, 0), (-1, 0), 2, self.colors['secondary_blue']),

            # Espaciado mejorado
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))

        story.append(table)
        story.append(Spacer(1, 30))

    def _add_correos_processing_section(self, story: List, data: Dict):
        """Agrega una subsección específica para el procesamiento de archivos de Correos de Costa Rica SA."""

        # Obtener estadísticas de Correos
        stats = data.get('processing_stats', {})
        correos_processed = stats.get('correos_pdfs_processed', 0)
        correos_failed = stats.get('correos_pdfs_failed', 0)
        correos_matches = stats.get('correos_matches', 0)

        # Solo mostrar la sección si se procesaron archivos de Correos
        if correos_processed == 0 and correos_failed == 0:
            return

        # Título de la subsección
        story.append(
            Paragraph("📮 Procesamiento Especial: Correos de Costa Rica SA", self.styles['CorreosSectionTitle']))

        # Texto explicativo
        explanation = f"Subsección de Nargallo del Este S.A. - Facturas procesadas mediante análisis de PDF"
        story.append(Paragraph(explanation, self.styles['InfoText']))

        # Crear tabla de estadísticas de Correos
        correos_stats_data = [
            ['Métrica de Correos', 'Resultado'],
            ['PDFs procesados exitosamente', f"{correos_processed:,}"],
            ['PDFs con errores de procesamiento', f"{correos_failed:,}"],
            ['Matches generados desde PDFs', f"{correos_matches:,}"],
            ['Tasa de éxito', f"{(correos_processed / (correos_processed + correos_failed) * 100):.1f}%" if (
                                                                                                                        correos_processed + correos_failed) > 0 else "N/A"]
        ]

        # Crear tabla de estadísticas de Correos
        correos_stats_table = Table(correos_stats_data, colWidths=[3.5 * inch, 2.0 * inch])

        correos_stats_table.setStyle(TableStyle([
            # Header con color verde para Correos
            ('BACKGROUND', (0, 0), (-1, 0), self.colors['correos_green']),
            ('TEXTCOLOR', (0, 0), (-1, 0), self.colors['white']),
            ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 11),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('TOPPADDING', (0, 0), (-1, 0), 12),

            # Body
            ('BACKGROUND', (0, 1), (-1, -1), self.colors['white']),
            ('TEXTCOLOR', (0, 1), (-1, -1), self.colors['text_dark']),
            ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
            ('FONTSIZE', (0, 1), (-1, -1), 10),
            ('BOTTOMPADDING', (0, 1), (-1, -1), 10),
            ('TOPPADDING', (0, 1), (-1, -1), 10),
            ('LEFTPADDING', (0, 0), (-1, -1), 15),
            ('RIGHTPADDING', (0, 0), (-1, -1), 15),

            # Alineación
            ('ALIGN', (0, 1), (0, -1), 'LEFT'),
            ('ALIGN', (1, 1), (-1, -1), 'CENTER'),

            # Filas alternadas con tema verde
            ('ROWBACKGROUNDS', (0, 1), (-1, -1), [self.colors['white'], self.colors['light_green']]),

            # Bordes
            ('GRID', (0, 0), (-1, -1), 1, self.colors['border_light']),
            ('LINEBELOW', (0, 0), (-1, 0), 2, self.colors['correos_green']),
            ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ]))

        story.append(correos_stats_table)

        # Agregar información adicional sobre el procesamiento de Correos
        if correos_processed > 0:
            story.append(Spacer(1, 15))

            info_text = """<b>Información del Procesamiento de Correos:</b><br/>
• Las facturas de Correos de Costa Rica SA requieren procesamiento especial<br/>
• Los datos se extraen directamente de los archivos PDF asociados<br/>
• Se extraen números de factura y códigos de guías automáticamente<br/>
• El formato de salida incluye: (Número Factura) SERVICIO EMS/GUIA [Códigos]<br/>
• Estos archivos forman parte del procesamiento de Nargallo del Este S.A."""

            story.append(Paragraph(info_text, self.styles['InfoText']))

        story.append(Spacer(1, 30))

    def _add_excluded_xmls_section(self, story: List, data: Dict):
        """Agrega una sección optimizada con los XMLs excluidos organizados por empresa."""

        # Obtener datos de XMLs excluidos
        excluded_details = data.get('excluded_xmls', [])

        if not excluded_details:
            return  # No mostrar sección si no hay XMLs excluidos

        # Título de la sección con color distintivo
        current_date = datetime.now()
        section_title = f"XMLs No Procesados (Fuera del mes actual: {current_date.month:02d}/{current_date.year})"
        story.append(Paragraph(section_title, self.styles['SecondarySectionTitle']))

        # Texto explicativo
        explanation = f"Total de XMLs excluidos del análisis: {len(excluded_details):,}"
        story.append(Paragraph(explanation, self.styles['InfoText']))

        # Organizar XMLs por empresa
        xmls_by_company = defaultdict(list)
        for excluded in excluded_details:
            company = excluded.get('company', 'Desconocida')
            xmls_by_company[company].append(excluded)

        # Orden de empresas predefinido
        company_order = ['Su Laka', 'Nargallo del Este S.A.', 'Ventas Fruno, S.A.', 'Creme Caramel']

        # Procesar cada empresa
        for company_name in company_order:
            if company_name not in xmls_by_company:
                continue

            company_xmls = xmls_by_company[company_name]

            # Título de la empresa
            story.append(Paragraph(f"🏢 {company_name} ({len(company_xmls):,} XMLs)", self.styles['CompanyTitle']))

            # Preparar datos para la tabla de esta empresa
            table_data = [['Número Consecutivo', 'Fecha de Emisión']]

            # Ordenar por fecha
            company_xmls.sort(key=lambda x: x.get('fecha_parsed', ''))

            for xml in company_xmls:
                numero = xml.get('numero_consecutivo', 'N/A')
                fecha = xml.get('fecha_parsed', 'N/A')

                # Truncar número si es muy largo
                if len(numero) > 25:
                    numero = numero[:22] + "..."

                table_data.append([numero, fecha])

            # Crear tabla para esta empresa
            company_table = Table(table_data, colWidths=[4.5 * inch, 2.0 * inch])

            company_table.setStyle(TableStyle([
                # Header
                ('BACKGROUND', (0, 0), (-1, 0), self.colors['warning_orange']),
                ('TEXTCOLOR', (0, 0), (-1, 0), self.colors['white']),
                ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                ('FONTSIZE', (0, 0), (-1, 0), 10),
                ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
                ('TOPPADDING', (0, 0), (-1, 0), 10),

                # Body
                ('BACKGROUND', (0, 1), (-1, -1), self.colors['white']),
                ('TEXTCOLOR', (0, 1), (-1, -1), self.colors['text_dark']),
                ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                ('FONTSIZE', (0, 1), (-1, -1), 9),
                ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
                ('TOPPADDING', (0, 1), (-1, -1), 8),
                ('LEFTPADDING', (0, 0), (-1, -1), 10),
                ('RIGHTPADDING', (0, 0), (-1, -1), 10),

                # Alineación
                ('ALIGN', (0, 1), (0, -1), 'LEFT'),  # Números a la izquierda
                ('ALIGN', (1, 1), (-1, -1), 'CENTER'),  # Fechas centradas

                # Filas alternadas
                ('ROWBACKGROUNDS', (0, 1), (-1, -1), [self.colors['white'], self.colors['light_orange']]),

                # Bordes
                ('GRID', (0, 0), (-1, -1), 1, self.colors['border_light']),
                ('LINEBELOW', (0, 0), (-1, 0), 2, self.colors['warning_orange']),
                ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
            ]))

            story.append(company_table)
            story.append(Spacer(1, 20))

        # Procesar empresas no predefinidas
        for company, xmls in xmls_by_company.items():
            if company not in company_order:
                story.append(Paragraph(f"🏢 {company} ({len(xmls):,} XMLs)", self.styles['CompanyTitle']))

                table_data = [['Número Consecutivo', 'Fecha de Emisión']]

                xmls.sort(key=lambda x: x.get('fecha_parsed', ''))

                for xml in xmls:
                    numero = xml.get('numero_consecutivo', 'N/A')
                    fecha = xml.get('fecha_parsed', 'N/A')

                    if len(numero) > 25:
                        numero = numero[:22] + "..."

                    table_data.append([numero, fecha])

                company_table = Table(table_data, colWidths=[4.5 * inch, 2.0 * inch])

                company_table.setStyle(TableStyle([
                    ('BACKGROUND', (0, 0), (-1, 0), self.colors['warning_orange']),
                    ('TEXTCOLOR', (0, 0), (-1, 0), self.colors['white']),
                    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
                    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
                    ('FONTSIZE', (0, 0), (-1, 0), 10),
                    ('BOTTOMPADDING', (0, 0), (-1, 0), 10),
                    ('TOPPADDING', (0, 0), (-1, 0), 10),
                    ('BACKGROUND', (0, 1), (-1, -1), self.colors['white']),
                    ('TEXTCOLOR', (0, 1), (-1, -1), self.colors['text_dark']),
                    ('FONTNAME', (0, 1), (-1, -1), 'Helvetica'),
                    ('FONTSIZE', (0, 1), (-1, -1), 9),
                    ('BOTTOMPADDING', (0, 1), (-1, -1), 8),
                    ('TOPPADDING', (0, 1), (-1, -1), 8),
                    ('LEFTPADDING', (0, 0), (-1, -1), 10),
                    ('RIGHTPADDING', (0, 0), (-1, -1), 10),
                    ('ALIGN', (0, 1), (0, -1), 'LEFT'),
                    ('ALIGN', (1, 1), (-1, -1), 'CENTER'),
                    ('ROWBACKGROUNDS', (0, 1), (-1, -1), [self.colors['white'], self.colors['light_orange']]),
                    ('GRID', (0, 0), (-1, -1), 1, self.colors['border_light']),
                    ('LINEBELOW', (0, 0), (-1, 0), 2, self.colors['warning_orange']),
                    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
                ]))

                story.append(company_table)
                story.append(Spacer(1, 20))

    def _add_clean_footer(self, story: List):
        """Agrega un footer limpio y discreto."""
        story.append(Spacer(1, 50))

        # Separador minimalista
        story.append(HRFlowable(width="40%", thickness=1,
                                color=self.colors['border_light'], spaceBefore=0, spaceAfter=20))

        # Footer con información esencial
        footer_text = f"""<b>ContaFlow</b> • Sistema de Automatización de Procesamiento de Datos Empresariales<br/>
Reporte generado automáticamente el {datetime.now().strftime('%d/%m/%Y a las %H:%M')}<br/>
Procesamiento empresarial optimizado con filtrado automático por fecha del mes actual<br/>
Incluye procesamiento especial para facturas de Correos de Costa Rica SA"""

        story.append(Paragraph(footer_text, self.styles['CleanFooter']))

    def generate_simple_summary_pdf(self, files: List[Dict], output_path: str,
                                    excluded_xmls: List[Dict] = None) -> bool:
        """
        Genera un resumen de texto simple cuando reportlab no está disponible, incluyendo XMLs excluidos y Correos.

        Args:
            files (List[Dict]): Lista de archivos procesados
            output_path (str): Ruta donde guardar el archivo
            excluded_xmls (List[Dict]): Lista de XMLs excluidos por fecha

        Returns:
            bool: True si se generó exitosamente
        """
        try:
            os.makedirs(os.path.dirname(output_path), exist_ok=True)
            text_output = output_path.replace('.pdf', '_resumen.txt')

            current_date = datetime.now()

            with open(text_output, 'w', encoding='utf-8') as f:
                f.write("FRUNO - CONTAFLOW\n")
                f.write("REPORTE DE CONTABILIDAD\n")
                f.write("=" * 50 + "\n\n")
                f.write(f"Generado el: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Mes actual procesado: {current_date.month:02d}/{current_date.year}\n")
                f.write("Procesado por ContaFlow\n\n")

                f.write("RESUMEN POR EMPRESAS:\n")
                f.write("-" * 25 + "\n")

                # Agrupar por empresas
                company_stats = {}
                for file_info in files:
                    company = file_info.get('company_name', 'Empresa Desconocida')
                    if company not in company_stats:
                        company_stats[company] = {'files': 0, 'matches': 0, 'manual_reviews': 0}

                    company_stats[company]['files'] += 1
                    company_stats[company]['matches'] += file_info.get('matches', 0)
                    company_stats[company]['manual_reviews'] += file_info.get('manual_reviews', 0)

                for company, stats in company_stats.items():
                    f.write(f"\n🏢 {company}\n")
                    f.write(f"   📁 Archivos: {stats['files']}\n")
                    f.write(f"   ✅ Matches: {stats['matches']:,}\n")
                    f.write(f"   🔍 Revision Manual: {stats['manual_reviews']}\n")

                # Sección especial de Correos (versión texto)
                f.write(f"\n\n📮 PROCESAMIENTO ESPECIAL - CORREOS DE COSTA RICA SA:\n")
                f.write("-" * 55 + "\n")
                f.write("Subsección de Nargallo del Este S.A.\n")
                f.write("Facturas procesadas mediante análisis de PDF\n\n")

                # Sección de XMLs excluidos organizados por empresa
                if excluded_xmls:
                    f.write(f"\n\nXMLS NO PROCESADOS (FUERA DEL MES ACTUAL):\n")
                    f.write("-" * 45 + "\n")
                    f.write(f"Total excluidos: {len(excluded_xmls)}\n\n")

                    # Organizar por empresa
                    xmls_by_company = defaultdict(list)
                    for excluded in excluded_xmls:
                        company = excluded.get('company', 'Desconocida')
                        xmls_by_company[company].append(excluded)

                    for company, xmls in xmls_by_company.items():
                        f.write(f"\n{company} ({len(xmls)} XMLs):\n")
                        for xml in xmls:
                            numero = xml.get('numero_consecutivo', 'N/A')
                            fecha = xml.get('fecha_parsed', 'N/A')
                            f.write(f"  - {numero} ({fecha})\n")

                f.write("\n" + "=" * 50 + "\n")
                f.write("ContaFlow - Sistema de Automatización Empresarial\n")
                f.write("Filtrado automático por fecha del mes actual\n")
                f.write("Procesamiento especial para Correos de Costa Rica SA\n")

            return True

        except Exception as e:
            print(f"Error generando resumen de texto: {e}")
            return False


def create_pdf_generator() -> Optional[PDFGenerator]:
    """
    Crea una instancia del generador de PDF si está disponible.

    Returns:
        PDFGenerator: Instancia del generador o None si no está disponible
    """
    try:
        return PDFGenerator()
    except ImportError:
        print("Advertencia: reportlab no está disponible. PDF de resumen no se generará.")
        print("Para habilitar PDF, instale reportlab: pip install reportlab")
        return None


def generate_processing_summary_pdf(processed_files: List[Dict], output_path: str,
                                    custom_message: str = None, processing_stats: Dict = None,
                                    excluded_xmls: List[Dict] = None) -> bool:
    """
    Función de utilidad para generar PDF limpio y profesional con información de XMLs excluidos y Correos.

    Args:
        processed_files (List[Dict]): Lista de archivos procesados
        output_path (str): Ruta donde guardar el PDF
        custom_message (str): Mensaje adicional opcional
        processing_stats (Dict): Estadísticas del procesamiento incluyendo Correos
        excluded_xmls (List[Dict]): Lista de XMLs excluidos por fecha

    Returns:
        bool: True si se generó exitosamente
    """
    generator = create_pdf_generator()

    if generator:
        # Preparar datos para el PDF
        data = {
            'title': 'Reporte de Contabilidad',
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'files': processed_files,
            'custom_message': custom_message or "",
            'companies_processed': list(set(f.get('company_name', 'Empresa Desconocida') for f in processed_files)),
            'processing_stats': processing_stats or {},
            'excluded_xmls': excluded_xmls or []
        }

        return generator.generate_summary_pdf(data, output_path)
    else:
        # Fallback a texto simple
        temp_generator = PDFGenerator.__new__(PDFGenerator)
        return temp_generator.generate_simple_summary_pdf(processed_files, output_path, excluded_xmls)