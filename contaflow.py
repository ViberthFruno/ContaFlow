#!/usr/bin/env python3
"""
Script de entrada para ContaFlow.
Configura el PYTHONPATH y ejecuta la aplicación.
"""

import sys
import os

# Agregar el directorio src/ al PYTHONPATH
src_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src')
if src_path not in sys.path:
    sys.path.insert(0, src_path)

# Importar y ejecutar el main
from contaflow.main import main

if __name__ == "__main__":
    main()
