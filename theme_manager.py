# theme_manager.py
"""
Sistema de temas moderno y optimizado para ContaFlow.
Diseño simple, elegante y de alto rendimiento.
"""

import tkinter as tk
from tkinter import ttk


class ModernTheme:
    """Tema moderno optimizado para ContaFlow - Simple y elegante."""

    # ========== PALETA DE COLORES ==========
    # Colores principales
    PRIMARY = "#2C3E50"          # Azul oscuro profesional
    PRIMARY_LIGHT = "#34495E"    # Azul oscuro claro
    SECONDARY = "#3498DB"        # Azul brillante
    ACCENT = "#9B59B6"           # Morado elegante

    # Colores de estado
    SUCCESS = "#27AE60"          # Verde moderno
    WARNING = "#F39C12"          # Naranja suave
    DANGER = "#E74C3C"           # Rojo moderno
    INFO = "#3498DB"             # Azul información

    # Colores de fondo
    BG_MAIN = "#ECF0F1"          # Gris claro principal
    BG_SURFACE = "#FFFFFF"       # Blanco superficie
    BG_DARK = "#2C3E50"          # Fondo oscuro
    BG_DARKER = "#1A252F"        # Fondo más oscuro

    # Colores de texto
    TEXT_PRIMARY = "#2C3E50"     # Texto principal oscuro
    TEXT_SECONDARY = "#7F8C8D"   # Texto secundario gris
    TEXT_LIGHT = "#ECF0F1"       # Texto claro
    TEXT_WHITE = "#FFFFFF"       # Texto blanco

    # Colores de borde
    BORDER = "#BDC3C7"           # Borde gris suave
    BORDER_LIGHT = "#E0E4E8"     # Borde más claro
    BORDER_DARK = "#95A5A6"      # Borde oscuro

    # ========== TIPOGRAFÍA ==========
    FONT_FAMILY = "Segoe UI"
    FONT_FAMILY_MONO = "Consolas"

    # Tamaños
    FONT_TITLE = ("Segoe UI", 24, "bold")
    FONT_HEADING = ("Segoe UI", 14, "bold")
    FONT_SUBHEADING = ("Segoe UI", 12, "bold")
    FONT_NORMAL = ("Segoe UI", 10)
    FONT_SMALL = ("Segoe UI", 9)
    FONT_MONO = ("Consolas", 9)
    FONT_MONO_NORMAL = ("Consolas", 10)

    # ========== ESPACIADO ==========
    PAD_SMALL = 5
    PAD_MEDIUM = 10
    PAD_LARGE = 15
    PAD_XLARGE = 20

    # ========== DIMENSIONES ==========
    BORDER_WIDTH = 1
    BUTTON_HEIGHT = 35
    INPUT_HEIGHT = 30

    @staticmethod
    def apply_theme(root):
        """
        Aplica el tema moderno a la aplicación de forma optimizada.

        Args:
            root: Ventana raíz de tkinter
        """
        style = ttk.Style(root)

        # Configurar tema base (optimizado para Windows)
        try:
            # Intentar usar tema moderno si está disponible
            available_themes = style.theme_names()
            if 'vista' in available_themes:
                style.theme_use('vista')
            elif 'clam' in available_themes:
                style.theme_use('clam')
            elif 'alt' in available_themes:
                style.theme_use('alt')
        except:
            pass

        # Aplicar estilos personalizados (en orden de uso)
        ModernTheme._configure_frames(style)
        ModernTheme._configure_labels(style)
        ModernTheme._configure_buttons(style)
        ModernTheme._configure_entries(style)
        ModernTheme._configure_notebook(style)
        ModernTheme._configure_labelframes(style)
        ModernTheme._configure_misc(style)

        # Configurar colores de fondo de la ventana principal
        root.configure(bg=ModernTheme.BG_MAIN)

        print("✨ Tema moderno aplicado con éxito")

    @staticmethod
    def _configure_frames(style):
        """Configura estilos de frames (optimizado)."""
        # Frame principal
        style.configure("TFrame",
                       background=ModernTheme.BG_MAIN)

        # Frame de superficie (cards)
        style.configure("Surface.TFrame",
                       background=ModernTheme.BG_SURFACE,
                       relief="flat")

    @staticmethod
    def _configure_labels(style):
        """Configura estilos de labels (optimizado)."""
        # Label normal
        style.configure("TLabel",
                       background=ModernTheme.BG_MAIN,
                       foreground=ModernTheme.TEXT_PRIMARY,
                       font=ModernTheme.FONT_NORMAL)

        # Label de título
        style.configure("Title.TLabel",
                       background=ModernTheme.BG_MAIN,
                       foreground=ModernTheme.PRIMARY,
                       font=ModernTheme.FONT_TITLE)

        # Label de heading
        style.configure("Heading.TLabel",
                       background=ModernTheme.BG_MAIN,
                       foreground=ModernTheme.TEXT_PRIMARY,
                       font=ModernTheme.FONT_HEADING)

        # Label de estado
        style.configure("Status.TLabel",
                       background=ModernTheme.BG_SURFACE,
                       foreground=ModernTheme.TEXT_PRIMARY,
                       font=ModernTheme.FONT_NORMAL,
                       padding=(10, 5))

    @staticmethod
    def _configure_buttons(style):
        """Configura estilos de botones con efectos hover (optimizado)."""
        # Botón principal (Primary)
        style.configure("Primary.TButton",
                       background=ModernTheme.SECONDARY,
                       foreground=ModernTheme.TEXT_WHITE,
                       font=ModernTheme.FONT_NORMAL,
                       borderwidth=0,
                       focuscolor='none',
                       padding=(15, 8))

        style.map("Primary.TButton",
                 background=[('active', '#2980B9'), ('pressed', '#2471A3'), ('disabled', '#95A5A6')],
                 foreground=[('active', ModernTheme.TEXT_WHITE), ('pressed', ModernTheme.TEXT_WHITE), ('disabled', '#ECF0F1')])

        # Botón de éxito (Success)
        style.configure("Success.TButton",
                       background=ModernTheme.SUCCESS,
                       foreground=ModernTheme.TEXT_WHITE,
                       font=ModernTheme.FONT_NORMAL,
                       borderwidth=0,
                       focuscolor='none',
                       padding=(15, 8))

        style.map("Success.TButton",
                 background=[('active', '#229954'), ('pressed', '#1E8449'), ('disabled', '#95A5A6')],
                 foreground=[('active', ModernTheme.TEXT_WHITE), ('pressed', ModernTheme.TEXT_WHITE), ('disabled', '#ECF0F1')])

        # Botón de peligro (Danger)
        style.configure("Danger.TButton",
                       background=ModernTheme.DANGER,
                       foreground=ModernTheme.TEXT_WHITE,
                       font=ModernTheme.FONT_NORMAL,
                       borderwidth=0,
                       focuscolor='none',
                       padding=(15, 8))

        style.map("Danger.TButton",
                 background=[('active', '#C0392B'), ('pressed', '#A93226'), ('disabled', '#95A5A6')],
                 foreground=[('active', ModernTheme.TEXT_WHITE), ('pressed', ModernTheme.TEXT_WHITE), ('disabled', '#ECF0F1')])

        # Botón normal (default)
        style.configure("TButton",
                       background=ModernTheme.BG_SURFACE,
                       foreground=ModernTheme.TEXT_PRIMARY,
                       font=ModernTheme.FONT_NORMAL,
                       borderwidth=1,
                       focuscolor='none',
                       padding=(12, 6))

        style.map("TButton",
                 background=[('active', ModernTheme.BORDER_LIGHT), ('pressed', ModernTheme.SECONDARY), ('disabled', '#E0E4E8')],
                 foreground=[('pressed', ModernTheme.TEXT_WHITE), ('disabled', '#95A5A6')],
                 bordercolor=[('active', ModernTheme.SECONDARY)])

        # Botón pequeño
        style.configure("Small.TButton",
                       background=ModernTheme.BG_SURFACE,
                       foreground=ModernTheme.TEXT_PRIMARY,
                       font=ModernTheme.FONT_SMALL,
                       borderwidth=1,
                       focuscolor='none',
                       padding=(8, 4))

        style.map("Small.TButton",
                 background=[('active', ModernTheme.BORDER_LIGHT), ('pressed', ModernTheme.SECONDARY), ('disabled', '#E0E4E8')],
                 foreground=[('pressed', ModernTheme.TEXT_WHITE), ('disabled', '#95A5A6')],
                 bordercolor=[('active', ModernTheme.SECONDARY)])

    @staticmethod
    def _configure_entries(style):
        """Configura estilos de entradas (optimizado)."""
        style.configure("TEntry",
                       fieldbackground=ModernTheme.BG_SURFACE,
                       foreground=ModernTheme.TEXT_PRIMARY,
                       borderwidth=1,
                       relief="solid",
                       padding=(8, 6))

        style.map("TEntry",
                 bordercolor=[('focus', ModernTheme.SECONDARY)],
                 lightcolor=[('focus', ModernTheme.SECONDARY)],
                 darkcolor=[('focus', ModernTheme.SECONDARY)])

        # Combobox
        style.configure("TCombobox",
                       fieldbackground=ModernTheme.BG_SURFACE,
                       foreground=ModernTheme.TEXT_PRIMARY,
                       background=ModernTheme.BG_SURFACE,
                       borderwidth=1,
                       padding=(8, 6))

        style.map("TCombobox",
                 fieldbackground=[('readonly', ModernTheme.BG_SURFACE)],
                 bordercolor=[('focus', ModernTheme.SECONDARY)])

    @staticmethod
    def _configure_notebook(style):
        """Configura estilos de notebook/pestañas (optimizado)."""
        # Notebook
        style.configure("TNotebook",
                       background=ModernTheme.BG_MAIN,
                       borderwidth=0,
                       tabmargins=[2, 5, 2, 0])

        style.configure("TNotebook.Tab",
                       background=ModernTheme.BG_SURFACE,
                       foreground=ModernTheme.TEXT_SECONDARY,
                       font=ModernTheme.FONT_NORMAL,
                       padding=[20, 10],
                       borderwidth=0)

        style.map("TNotebook.Tab",
                 background=[('selected', ModernTheme.SECONDARY), ('active', ModernTheme.BORDER_LIGHT), ('!selected', ModernTheme.BG_SURFACE)],
                 foreground=[('selected', ModernTheme.TEXT_WHITE), ('active', ModernTheme.PRIMARY), ('!selected', ModernTheme.TEXT_SECONDARY)],
                 expand=[('selected', [1, 1, 1, 0])])

    @staticmethod
    def _configure_labelframes(style):
        """Configura estilos de labelframes (optimizado)."""
        style.configure("TLabelframe",
                       background=ModernTheme.BG_SURFACE,
                       foreground=ModernTheme.TEXT_PRIMARY,
                       borderwidth=1,
                       relief="solid",
                       bordercolor=ModernTheme.BORDER_LIGHT)

        style.configure("TLabelframe.Label",
                       background=ModernTheme.BG_SURFACE,
                       foreground=ModernTheme.PRIMARY,
                       font=ModernTheme.FONT_SUBHEADING)

        # LabelFrame moderno con sombra visual
        style.configure("Modern.TLabelframe",
                       background=ModernTheme.BG_SURFACE,
                       foreground=ModernTheme.TEXT_PRIMARY,
                       borderwidth=0,
                       relief="flat")

        style.configure("Modern.TLabelframe.Label",
                       background=ModernTheme.BG_SURFACE,
                       foreground=ModernTheme.PRIMARY,
                       font=ModernTheme.FONT_HEADING)

    @staticmethod
    def _configure_misc(style):
        """Configura otros widgets (optimizado)."""
        # Separator
        style.configure("TSeparator",
                       background=ModernTheme.BORDER_LIGHT)

        # Scrollbar
        style.configure("TScrollbar",
                       background=ModernTheme.BG_SURFACE,
                       troughcolor=ModernTheme.BG_MAIN,
                       borderwidth=0,
                       arrowcolor=ModernTheme.TEXT_SECONDARY)

        style.map("TScrollbar",
                 background=[('active', ModernTheme.BORDER)])

        # Progressbar
        style.configure("TProgressbar",
                       background=ModernTheme.SUCCESS,
                       troughcolor=ModernTheme.BG_MAIN,
                       borderwidth=0,
                       thickness=20)

    @staticmethod
    def get_button_style(button_type="normal"):
        """
        Obtiene el estilo de botón apropiado (método de ayuda rápido).

        Args:
            button_type: 'primary', 'success', 'danger', 'normal', 'small'

        Returns:
            str: Nombre del estilo ttk
        """
        styles = {
            'primary': 'Primary.TButton',
            'success': 'Success.TButton',
            'danger': 'Danger.TButton',
            'normal': 'TButton',
            'small': 'Small.TButton'
        }
        return styles.get(button_type, 'TButton')

    @staticmethod
    def create_status_badge(parent, text, status_type="info"):
        """
        Crea un badge de estado moderno (optimizado).

        Args:
            parent: Widget padre
            text: Texto del badge
            status_type: 'success', 'warning', 'danger', 'info'

        Returns:
            ttk.Label: Badge configurado
        """
        colors = {
            'success': (ModernTheme.SUCCESS, ModernTheme.TEXT_WHITE),
            'warning': (ModernTheme.WARNING, ModernTheme.TEXT_WHITE),
            'danger': (ModernTheme.DANGER, ModernTheme.TEXT_WHITE),
            'info': (ModernTheme.INFO, ModernTheme.TEXT_WHITE)
        }

        bg, fg = colors.get(status_type, colors['info'])

        badge = tk.Label(parent,
                        text=text,
                        bg=bg,
                        fg=fg,
                        font=ModernTheme.FONT_SMALL,
                        padx=10,
                        pady=4,
                        relief="flat")

        return badge


# ========== FUNCIONES DE AYUDA RÁPIDA ==========

def apply_modern_theme(root):
    """
    Aplica el tema moderno a la aplicación (función de ayuda).

    Args:
        root: Ventana raíz de tkinter
    """
    ModernTheme.apply_theme(root)


def get_color(color_name):
    """
    Obtiene un color del tema (función de ayuda rápida).

    Args:
        color_name: Nombre del color en mayúsculas

    Returns:
        str: Código de color hexadecimal
    """
    return getattr(ModernTheme, color_name, ModernTheme.PRIMARY)


def create_modern_text_widget(parent, **kwargs):
    """
    Crea un widget Text con estilo moderno (optimizado).

    Args:
        parent: Widget padre
        **kwargs: Argumentos adicionales para Text

    Returns:
        tk.Text: Widget de texto configurado
    """
    defaults = {
        'bg': ModernTheme.BG_DARK,
        'fg': ModernTheme.TEXT_LIGHT,
        'font': ModernTheme.FONT_MONO,
        'relief': 'flat',
        'borderwidth': 0,
        'padx': 10,
        'pady': 10,
        'insertbackground': ModernTheme.TEXT_LIGHT,
        'selectbackground': ModernTheme.SECONDARY,
        'selectforeground': ModernTheme.TEXT_WHITE
    }

    # Combinar defaults con kwargs
    config = {**defaults, **kwargs}

    return tk.Text(parent, **config)
