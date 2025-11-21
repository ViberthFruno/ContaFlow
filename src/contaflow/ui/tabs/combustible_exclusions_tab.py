"""combustible_exclusions_tab.py

Pestaña de configuración para gestionar exclusiones de combustibles/placas
basadas en el campo <NombreEmisor>.
Diseño moderno optimizado.
"""

import tkinter as tk
from tkinter import ttk
from typing import List
import unicodedata

from contaflow.config.config_manager import ConfigManager
from contaflow.ui.theme_manager import ModernTheme


class CombustibleExclusionsTab:
    """Gestiona las exclusiones de extracción de placas por NombreEmisor."""

    def __init__(self, parent):
        self.parent = parent
        self.config_manager = ConfigManager()
        self.is_visible = False

        self.emitter_var = tk.StringVar()
        self.exclusions: List[str] = []
        self.xml_config_available = False

        self.listbox = None
        self.status_label = None

        self._create_interface()
        self.load_exclusions()

    def _create_interface(self):
        main_frame = ttk.Frame(self.parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(2, weight=1)

        # Descripción moderna con card
        desc_card = tk.Frame(main_frame, bg=ModernTheme.INFO,
                            highlightbackground=ModernTheme.SECONDARY,
                            highlightthickness=1)
        desc_card.grid(row=0, column=0, sticky="ew", pady=(0, 10))

        description = (
            "💬 Agrega los valores exactos de <NombreEmisor> que deben excluirse del "
            "tratamiento como Combustible/Placa. Cuando se detecte un emisor en "
            "esta lista, el bot utilizará el comportamiento normal (sin tratarlo "
            "como combustible)."
        )
        tk.Label(desc_card, text=description, wraplength=520, justify=tk.LEFT,
                bg=ModernTheme.INFO, fg=ModernTheme.TEXT_WHITE,
                font=ModernTheme.FONT_SMALL, pady=8, padx=10).pack()

        input_frame = ttk.LabelFrame(main_frame, text="➕ Agregar exclusión",
                                     padding=10, style="Modern.TLabelframe")
        input_frame.grid(row=1, column=0, sticky="ew")
        input_frame.columnconfigure(0, weight=1)

        entry = ttk.Entry(input_frame, textvariable=self.emitter_var,
                         font=ModernTheme.FONT_NORMAL)
        entry.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        add_btn = ttk.Button(input_frame, text="Agregar", command=self.add_exclusion,
                            style="Primary.TButton")
        add_btn.grid(row=0, column=1, ipady=4)

        list_frame = ttk.LabelFrame(main_frame, text="📋 Lista de emisores excluidos",
                                   padding=10, style="Modern.TLabelframe")
        list_frame.grid(row=2, column=0, sticky="nsew", pady=10)
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        self.listbox = tk.Listbox(list_frame, height=10,
                                  font=ModernTheme.FONT_NORMAL,
                                  bg=ModernTheme.BG_SURFACE,
                                  fg=ModernTheme.TEXT_PRIMARY,
                                  selectbackground=ModernTheme.SECONDARY,
                                  selectforeground=ModernTheme.TEXT_WHITE)
        self.listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.listbox.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.listbox.configure(yscrollcommand=scrollbar.set)

        buttons_frame = ttk.Frame(main_frame)
        buttons_frame.grid(row=3, column=0, sticky="ew")
        buttons_frame.columnconfigure((0, 1, 2), weight=1)

        remove_btn = ttk.Button(buttons_frame, text="Eliminar seleccionado",
                               command=self.remove_selected, style="TButton")
        remove_btn.grid(row=0, column=0, sticky="ew", padx=(0, 5), ipady=5)

        clear_btn = ttk.Button(buttons_frame, text="Limpiar lista",
                              command=self.clear_exclusions, style="TButton")
        clear_btn.grid(row=0, column=1, sticky="ew", padx=5, ipady=5)

        save_btn = ttk.Button(buttons_frame, text="Guardar cambios",
                             command=self.save_exclusions, style="Success.TButton")
        save_btn.grid(row=0, column=2, sticky="ew", padx=(5, 0), ipady=5)

        status_frame = ttk.LabelFrame(main_frame, text="Estado", padding=10)
        status_frame.grid(row=4, column=0, sticky="ew", pady=(10, 0))

        self.status_label = ttk.Label(status_frame, text="🔴 Sin configurar", foreground="red")
        self.status_label.pack(fill=tk.X)

    def load_exclusions(self):
        try:
            config = self.config_manager.load_config() or {}
            xml_config = config.get('xml_config')

            if isinstance(xml_config, dict) and (
                xml_config.get('company_folders') or xml_config.get('xml_folder')
            ):
                self.xml_config_available = True
                exclusions_config = xml_config.get('combustible_exclusions', {})

                if isinstance(exclusions_config, dict):
                    emitter_names = exclusions_config.get('emitter_names', [])
                elif isinstance(exclusions_config, list):
                    emitter_names = exclusions_config
                else:
                    emitter_names = []

                self.exclusions = [name for name in emitter_names if isinstance(name, str) and name.strip()]
                self.exclusions.sort(key=lambda x: x.lower())
                self._refresh_listbox()

                if self.exclusions:
                    self._update_status("🟢 Exclusiones cargadas correctamente.", "green")
                else:
                    self._update_status("🟡 No hay exclusiones configuradas.", "orange")
            else:
                self.xml_config_available = False
                self.exclusions = []
                self._refresh_listbox()
                self._update_status(
                    "🔴 Debe configurar la pestaña XML antes de agregar exclusiones.",
                    "red"
                )
        except Exception as e:
            self._update_status(f"🔴 Error al cargar exclusiones: {e}", "red")

    def add_exclusion(self):
        name = self.emitter_var.get().strip()
        if not name:
            self._update_status("🔴 Debe ingresar un valor para <NombreEmisor>.", "red")
            return

        normalized = self._normalize_name(name)
        if normalized in {self._normalize_name(item) for item in self.exclusions}:
            self._update_status("⚠️ El emisor ya se encuentra en la lista.", "orange")
            return

        self.exclusions.append(name)
        self.exclusions.sort(key=lambda x: x.lower())
        self._refresh_listbox()
        self.emitter_var.set("")
        self._update_status("🟢 Emisor agregado a las exclusiones.", "green")

    def remove_selected(self):
        if not self.listbox.curselection():
            self._update_status("⚠️ Debe seleccionar un emisor para eliminar.", "orange")
            return

        index = self.listbox.curselection()[0]
        removed = self.exclusions.pop(index)
        self._refresh_listbox()
        self._update_status(f"🟢 Emisor eliminado: {removed}", "green")

    def clear_exclusions(self):
        if not self.exclusions:
            self._update_status("⚠️ La lista de exclusiones ya está vacía.", "orange")
            return

        self.exclusions.clear()
        self._refresh_listbox()
        self._update_status("🟡 Lista de exclusiones limpiada. Recuerde guardar los cambios.", "orange")

    def save_exclusions(self):
        if not self.xml_config_available:
            self._update_status("🔴 Configure la pestaña XML antes de guardar exclusiones.", "red")
            return

        try:
            config = self.config_manager.load_config() or {}
            xml_config = config.get('xml_config')

            if not isinstance(xml_config, dict):
                self._update_status("🔴 Configuración XML inválida.", "red")
                return

            updated_xml_config = xml_config.copy()
            updated_xml_config['combustible_exclusions'] = {
                'emitter_names': self.exclusions
            }

            self.config_manager.update_config({'xml_config': updated_xml_config})
            self._update_status("🟢 Exclusiones guardadas correctamente.", "green")
        except Exception as e:
            self._update_status(f"🔴 Error al guardar exclusiones: {e}", "red")

    def _refresh_listbox(self):
        if not self.listbox:
            return
        self.listbox.delete(0, tk.END)
        for item in self.exclusions:
            self.listbox.insert(tk.END, item)

    def _update_status(self, message: str, color: str):
        if self.status_label:
            self.status_label.config(text=message, foreground=color)

    @staticmethod
    def _normalize_name(name: str) -> str:
        if not name:
            return ""
        normalized = unicodedata.normalize('NFKD', name)
        normalized = ''.join(c for c in normalized if not unicodedata.combining(c))
        return normalized.strip().lower()

    def show(self):
        if not self.is_visible:
            self.is_visible = True

    def hide(self):
        if self.is_visible:
            self.is_visible = False
