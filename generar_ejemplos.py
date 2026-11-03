"""Crea únicamente datos ficticios para demostrar el programa."""
from pathlib import Path
from openpyxl import Workbook

def main():
    destino = Path(__file__).parent / 'ejemplos'
    destino.mkdir(exist_ok=True)
    datos = {
        'distribucion': [('001', 'Multas', 100), ('002', 'Multas', 120),
                         ('003', 'Multas', 50), ('004', 'Multas', 20),
                         ('004', 'Multas', 20), ('005', 'Multas', 'incorrecto')],
        'contabilidad': [('001', 'Multas', 100), ('002', 'Multas', 110),
                         ('004', 'Multas', 40), ('006', 'Multas', 30)]}
    for nombre, filas in datos.items():
        ruta = destino / f'{nombre}.xlsx'
        if ruta.exists():
            raise FileExistsError(f'No se sobrescribe {ruta}')
        libro = Workbook()
        libro.active.append(['identificacion', 'concepto', 'valor'])
        for fila in filas:
            libro.active.append(fila)
        libro.save(ruta)
    print('Ejemplos ficticios creados.')

if __name__ == '__main__':
    main()
