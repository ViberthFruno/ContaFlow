#!/usr/bin/env python3
"""
Script de entrada para ContaFlow.
Ejecuta la aplicación directamente.
"""

import sys
import os

# Agregar src/contaflow al path para que los imports funcionen
contaflow_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src', 'contaflow')
sys.path.insert(0, contaflow_path)

# Ejecutar el main directamente
if __name__ == "__main__":
    import runpy
    runpy.run_path(os.path.join(contaflow_path, 'main.py'), run_name='__main__')
