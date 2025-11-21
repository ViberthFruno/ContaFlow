# ContaFlow Bot
"""
Sistema automatizado de procesamiento de datos empresariales que descarga archivos Excel desde correos, los analiza con datos XML de carpetas compartidas y envía los resultados procesados.
"""

## 📋 Descripción

ContaFlow Bot es una aplicación de escritorio desarrollada para automatizar el procesamiento de datos empresariales mediante la descarga de archivos Excel desde correos electrónicos, su análisis con datos XML almacenados en carpetas compartidas y el envío automático de los resultados procesados.

Esta herramienta está diseñada para:

- 📥 **Descarga automática** de archivos Excel desde correos no leídos con criterios específicos
- 🔍 **Análisis de datos** con matching automático entre Excel y XMLs empresariales
- 🏢 **Procesamiento multi-empresa** con soporte para múltiples empresas y rutas dinámicas
- 📊 **Generación de reportes** con archivos Excel procesados listos para usar
- 📧 **Envío automático** de resultados consolidados por correo electrónico
- ⏰ **Automatización completa** con perfiles programables y ejecución automática
- 📅 **Filtrado por fecha** para procesar solo datos del mes actual

---

## ✨ Características Principales

| Característica | Descripción |
|---|---|
| 📥 **Descarga inteligente** | Busca y descarga archivos Excel de correos específicos con filtros configurables |
| 🔍 **Análisis de datos** | Procesa archivos Excel y los cruza con datos XML de carpetas compartidas |
| 🏢 **Multi-empresa** | Maneja múltiples empresas con carpetas y configuraciones independientes |
| 📅 **Filtrado temporal** | Procesa solo datos del mes actual, excluyendo registros obsoletos |
| 📮 **Soporte Correos CR** | Procesamiento especializado para facturas de Correos de Costa Rica SA |
| 📊 **Reportes automáticos** | Genera archivos Excel con matches encontrados y revisiones manuales |
| 📧 **Envío consolidado** | Envía todos los resultados en un solo correo con mensajes personalizables |
| 🤖 **Automatización total** | Ejecuta todo el proceso sin intervención manual |
| ⏰ **Programación flexible** | Configura horarios y días específicos para ejecución automática |
| 📱 **Interfaz intuitiva** | Panel de control con pestañas organizadas y logs en tiempo real |

---

## 🛠️ Instalación

### Opción 1: Instalación con pip (Recomendado)
1. Clone o descargue el repositorio del proyecto
2. Navegue al directorio del proyecto
3. Instale el paquete en modo desarrollo:
```bash
pip install -e .
```
4. Ejecute la aplicación:
```bash
contaflow
```

### Opción 2: Ejecución directa
1. Asegúrese de tener Python 3.7 o posterior
2. Clone o descargue el repositorio del proyecto
3. Instale las dependencias:
```bash
pip install openpyxl pandas pdfplumber pywin32
```
4. Ejecute la aplicación usando el script de entrada:
```bash
python contaflow.py
```

### Opción 3: Ejecutable compilado
1. Descargue `ContaFlow.exe` desde la ubicación interna compartida
2. Colóquelo en la ubicación deseada de su sistema
3. Ejecute la aplicación con doble clic

---

## 📘 Guía de Uso

### 1. Configuración Inicial
- **Configurar cuenta de correo**: Gmail, Outlook o Yahoo
- **Establecer criterios de búsqueda**: Asunto, tipo de archivos, fechas
- **Configurar carpetas XML**: Rutas de las carpetas compartidas por empresa
- **Definir destinatarios**: Correos que recibirán los resultados

### 2. Pestaña de Búsqueda
- Configure el asunto de los correos a buscar (ej: "Cargador")
- Establezca la carpeta de descarga
- Active filtros por fecha y tipo de archivo
- Guarde la configuración para uso futuro

### 3. Pestaña XML
- Configure las carpetas base de cada empresa
- Establezca actividades comerciales por empresa
- Configure límites de revisión manual
- Active/desactive funciones como eliminación de originales

### 4. Pestaña de Destinatarios
- Configure el destinatario principal
- Añada destinatarios en copia (CC)
- Personalice mensajes del correo
- Configure envío de PDF resumen

### 5. Automatización Completa
- Inicie el bot desde la pestaña principal
- Monitoree el progreso en tiempo real
- Revise estadísticas y logs detallados
- Detenga el proceso cuando sea necesario

---

## 🗂️ Estructura del Proyecto

```
ContaFlow/
├── contaflow.py              # Script de entrada para ejecutar la aplicación
├── setup.py                  # Configuración de instalación del paquete
├── config/                   # Archivos de configuración
│   └── contaflow_config.json
├── assets/                   # Recursos (iconos, imágenes)
│   └── icon.ico
└── src/
    └── contaflow/           # Paquete principal
        ├── main.py          # Punto de entrada de la aplicación
        ├── ui/              # Interfaz gráfica
        │   ├── main_window.py      # Ventana principal
        │   ├── theme_manager.py    # Gestión de temas
        │   └── tabs/               # Pestañas de la interfaz
        │       ├── automatizacion_tab.py
        │       ├── automatizacion_ui.py
        │       ├── configuracion_tab.py
        │       ├── busqueda_tab.py
        │       ├── xml_tab.py
        │       └── combustible_exclusions_tab.py
        ├── core/            # Lógica principal del negocio
        │   ├── bot_controller.py    # Controlador del bot
        │   └── excel_processor.py   # Procesamiento Excel/XML
        ├── email/           # Gestión de correos
        │   ├── email_manager.py     # Conexiones de correo
        │   ├── email_processor.py   # Procesamiento de emails
        │   └── email_sender.py      # Envío de correos
        ├── processors/      # Procesadores especializados
        │   ├── pdf_generator.py     # Generación de PDFs
        │   ├── pdf_processor.py     # Procesamiento de PDFs
        │   └── otro_texto_processor.py
        └── config/          # Gestión de configuración
            └── config_manager.py
```

### Módulos Principales

| Módulo | Descripción |
|---|---|
| **ui/** | Interfaz gráfica con Tkinter y gestión de pestañas |
| **core/** | Lógica principal: bot controller y procesamiento Excel/XML |
| **email/** | Gestión completa de correos: conexión, procesamiento y envío |
| **processors/** | Procesadores especializados para PDFs y otros formatos |
| **config/** | Gestor centralizado de configuraciones |

---

## 🔧 Características Técnicas

### Procesamiento de Datos
- **Rutas dinámicas automáticas** basadas en año/mes actual
- **Filtrado inteligente** que excluye datos fuera del mes corriente
- **Matching avanzado** entre archivos Excel y XMLs empresariales
- **Validación de datos** con detección de duplicados y errores
- **Procesamiento especializado** para diferentes tipos de documentos

### Sistema de Correos
- **Conexión robusta** con reintentos automáticos
- **Soporte multi-proveedor** (Gmail, Outlook, Yahoo)
- **Filtros avanzados** por asunto, fecha y tipo de archivo
- **Envío consolidado** con múltiples archivos adjuntos
- **Mensajes personalizables** con plantillas dinámicas

### Automatización
- **Ejecución programada** con validación de condiciones
- **Monitoreo en tiempo real** con logs detallados
- **Manejo de errores** con recuperación automática
- **Estadísticas completas** de procesamiento
- **Interfaz responsiva** que no bloquea durante operaciones

---

## 📊 Flujo de Trabajo

1. **Búsqueda**: El bot busca correos con criterios específicos
2. **Descarga**: Descarga archivos Excel adjuntos automáticamente
3. **Análisis**: Cruza datos del Excel con XMLs de carpetas compartidas
4. **Filtrado**: Excluye registros que no corresponden al mes actual
5. **Procesamiento**: Genera archivos por empresa con matches encontrados
6. **Consolidación**: Prepara archivos y estadísticas para envío
7. **Envío**: Envía todos los resultados en un correo consolidado
8. **Limpieza**: Elimina archivos temporales según configuración

---

## 🧯 Solución de Problemas

### 📧 Problemas de conexión de correo
- Verificar credenciales y configuración SMTP
- Comprobar conexión a internet y firewall
- Validar que la autenticación de aplicaciones esté habilitada

### 📊 Datos no procesados correctamente
- Verificar que las carpetas XML estén accesibles
- Comprobar permisos de lectura en carpetas compartidas
- Validar formato de archivos Excel descargados

### ⏱️ Problemas de rendimiento
- Verificar espacio disponible en disco
- Comprobar que no hay procesos bloqueantes
- Revisar logs para identificar cuellos de botella

### 🏢 Carpetas empresariales no encontradas
- Verificar rutas de carpetas compartidas
- Comprobar que existan subcarpetas de año/mes actual
- Validar permisos de acceso a carpetas de red

---

## ⚠️ Notas Importantes

- **Uso empresarial interno**: Diseñado específicamente para procesos internos
- **Dependencias de red**: Requiere acceso a carpetas compartidas y correo
- **Compatibilidad**: Optimizado para Windows con Microsoft Outlook
- **Seguridad**: Maneja credenciales encriptadas y datos empresariales sensibles
- **Rendimiento**: Diseñado para procesar grandes volúmenes de datos eficientemente
- **Mantenimiento**: Logs detallados facilitan diagnóstico y soporte técnico

---

## 📈 Beneficios

- **Ahorro de tiempo**: Automatiza tareas repetitivas que tomaban horas
- **Reducción de errores**: Elimina errores manuales en procesamiento de datos
- **Trazabilidad completa**: Logs detallados de todas las operaciones
- **Escalabilidad**: Maneja múltiples empresas y grandes volúmenes de datos
- **Flexibilidad**: Configuración adaptable a diferentes necesidades empresariales
- **Confiabilidad**: Sistema robusto con manejo de errores y recuperación automática

---

*Desarrollado para optimizar procesos empresariales y automatizar el análisis de datos contables*
