"""
Actualiza el cache local del sheet "Datos comprobantes".
Correr ANTES de imputar para tener los datos más recientes.

Uso: python exportar_comprobantes.py
"""

from comprobantes_helper import exportar_cache

if __name__ == '__main__':
    exportar_cache()
