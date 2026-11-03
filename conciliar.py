"""Conciliación reproducible de dos Excel. Python 3.10 o superior."""
import argparse
from collections import Counter, defaultdict
from decimal import Decimal, InvalidOperation
from pathlib import Path
import hashlib
import json
from datetime import datetime, timezone

from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill

REQUERIDAS = ('identificacion', 'concepto', 'valor')


def importe(valor):
    """Exige importes numéricos o texto decimal sin separadores de miles."""
    if valor is None or isinstance(valor, bool):
        raise ValueError('Importe vacío o inválido')
    try:
        numero = Decimal(str(valor).strip())
    except InvalidOperation as exc:
        raise ValueError('Use números o texto como 1234.56, sin símbolos ni miles') from exc
    if not numero.is_finite() or numero != numero.quantize(Decimal('0.01')):
        raise ValueError('El importe debe ser finito y tener máximo dos decimales')
    return int(numero * 100)


def leer(ruta, hoja=None):
    libro = load_workbook(ruta, read_only=True, data_only=False)
    try:
        tabla = libro[hoja] if hoja else libro.worksheets[0]
        filas = tabla.iter_rows(values_only=True)
        cabecera = next(filas, ())
        nombres = [str(c).strip().lower() if c is not None else '' for c in cabecera]
        if any(nombres.count(c) != 1 for c in REQUERIDAS):
            raise ValueError(f'{ruta}: se requieren columnas únicas {REQUERIDAS}')
        indices = [nombres.index(c) for c in REQUERIDAS]
        validos, invalidos = [], []
        for numero, fila in enumerate(filas, 2):
            if all(v is None for v in fila):
                continue
            datos = [fila[i] for i in indices]
            try:
                if any(isinstance(v, str) and v.startswith('=') for v in datos):
                    raise ValueError('No se admiten fórmulas en las columnas del cruce')
                if any(v is None or not str(v).strip() for v in datos[:2]):
                    raise ValueError('Identificación o concepto vacío')
                identificacion, concepto = [str(v).strip() for v in datos[:2]]
                validos.append((numero, identificacion, concepto, importe(datos[2])))
            except ValueError as exc:
                invalidos.append((str(ruta), tabla.title, numero, str(exc), *map(str, datos)))
        return validos, invalidos
    finally:
        libro.close()


def agrupar(registros):
    grupos = defaultdict(lambda: [0, []])
    for fila, identificacion, concepto, centavos in registros:
        grupo = grupos[(identificacion, concepto)]
        grupo[0] += centavos
        grupo[1].append(fila)
    return grupos


def cruzar(a, b, tolerancia=0):
    izquierda, derecha = agrupar(a), agrupar(b)
    resultado = []
    for clave in sorted(izquierda.keys() | derecha.keys()):
        va, fa = izquierda.get(clave, [0, []])
        vb, fb = derecha.get(clave, [0, []])
        estado = ('Solo distribución' if not fb else 'Solo contabilidad' if not fa
                  else 'Coincide' if abs(va-vb) <= tolerancia else 'Diferencia')
        resultado.append((*clave, va/100, vb/100, (va-vb)/100, estado,
                          ', '.join(map(str, fa)), ', '.join(map(str, fb))))
    return resultado


def escribir_hoja(libro, nombre, cabecera, filas):
    hoja = libro.create_sheet(nombre)
    hoja.append(cabecera)
    for fila in filas:
        hoja.append(list(fila))
    # Fuerza cadenas literales: los datos de entrada nunca son fórmulas.
    for fila in hoja:
        for celda in fila:
            if isinstance(celda.value, str):
                celda.data_type = 's'
    for celda in hoja[1]:
        celda.font = Font(color='FFFFFF', bold=True)
        celda.fill = PatternFill('solid', fgColor='17365D')
    hoja.freeze_panes = 'A2'
    hoja.auto_filter.ref = hoja.dimensions
    for columna in hoja.columns:
        letra = columna[0].column_letter
        hoja.column_dimensions[letra].width = min(65, max(16, max(len(str(c.value or '')) for c in columna)+2))
    return hoja


def ejecutar(distribucion, contabilidad, salida, hoja_d=None, hoja_c=None, tolerancia=0):
    rutas = [Path(p).resolve() for p in (distribucion, contabilidad, salida)]
    if rutas[2] in rutas[:2] or rutas[2].exists():
        raise ValueError('La salida debe ser un archivo nuevo y distinto de las entradas')
    if tolerancia < 0:
        raise ValueError('La tolerancia no puede ser negativa')
    a, errores_a = leer(rutas[0], hoja_d)
    b, errores_b = leer(rutas[1], hoja_c)
    if not a or not b:
        raise ValueError('Ambos archivos deben contener al menos un registro válido')
    cruce = cruzar(a, b, tolerancia)
    libro = Workbook()
    libro.remove(libro.active)
    errores = errores_a + errores_b
    resumen = [('Estado del proceso', 'PARCIAL: hay filas excluidas' if errores else 'Completo'),
               ('Filas válidas distribución', len(a)), ('Filas válidas contabilidad', len(b)),
               ('Filas inválidas excluidas', len(errores)),
               ('Total válido distribución', sum(r[3] for r in a)/100),
               ('Total válido contabilidad', sum(r[3] for r in b)/100),
               ('Diferencia neta válida D - C', (sum(r[3] for r in a)-sum(r[3] for r in b))/100),
               ('Tolerancia por grupo', tolerancia/100)]
    resumen.extend(Counter(r[5] for r in cruce).items())
    escribir_hoja(libro, 'Resumen', ['Indicador', 'Resultado'], resumen)
    hoja = escribir_hoja(libro, 'Cruce', ['identificacion', 'concepto', 'distribucion', 'contabilidad', 'diferencia_D_menos_C', 'estado', 'filas_distribucion', 'filas_contabilidad'], cruce)
    for fila in hoja.iter_rows(min_row=2, min_col=3, max_col=5):
        for celda in fila:
            celda.number_format = '#,##0.00'
    escribir_hoja(libro, 'Invalidos', ['archivo', 'hoja', 'fila', 'motivo', *REQUERIDAS], errores)
    duplicados = []
    for origen, registros in [('Distribución', a), ('Contabilidad', b)]:
        exactos = defaultdict(list)
        for fila, ident, concepto, valor in registros:
            exactos[(ident, concepto, valor)].append(fila)
        for (ident, concepto, valor), filas in exactos.items():
            if len(filas) > 1:
                duplicados.append((origen, ident, concepto, valor/100, len(filas), ', '.join(map(str, filas))))
    escribir_hoja(libro, 'Posibles_duplicados', ['origen', 'identificacion', 'concepto', 'valor', 'repeticiones', 'filas'], duplicados)
    for nombre, registros in [('Base_distribucion', a), ('Base_contabilidad', b)]:
        escribir_hoja(libro, nombre, ['fila_original', 'identificacion', 'concepto', 'valor'], [(f, i, c, v/100) for f, i, c, v in registros])
    metadata = [('fecha_UTC', datetime.now(timezone.utc).isoformat()),
                ('regla', 'Agrupar identificación y concepto; sumar registros válidos sin eliminar repeticiones'),
                ('alcance', 'Una moneda y periodo comparables, verificados por el usuario'),
                ('advertencia', 'Diferencias y repeticiones requieren revisión; no prueban pérdida ni fraude')]
    for nombre, ruta in zip(('distribucion', 'contabilidad'), rutas[:2]):
        metadata.extend([(nombre, str(ruta)), (nombre+'_sha256', hashlib.sha256(ruta.read_bytes()).hexdigest())])
    escribir_hoja(libro, 'Trazabilidad', ['Dato', 'Valor'], metadata)
    rutas[2].parent.mkdir(parents=True, exist_ok=True)
    libro.save(rutas[2])
    return dict(resumen)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('distribucion')
    parser.add_argument('contabilidad')
    parser.add_argument('--salida', default='informe_conciliacion.xlsx')
    parser.add_argument('--hoja-distribucion')
    parser.add_argument('--hoja-contabilidad')
    parser.add_argument('--tolerancia', default='0.00', help='Importe decimal por grupo; predeterminado 0.00')
    args = parser.parse_args()
    try:
        resumen = ejecutar(args.distribucion, args.contabilidad, args.salida,
                           args.hoja_distribucion, args.hoja_contabilidad, importe(args.tolerancia))
    except (ValueError, OSError, KeyError, InvalidOperation) as exc:
        parser.exit(1, f'Error: {exc}\n')
    print(json.dumps(resumen, ensure_ascii=False, indent=2))
    print(f'Informe creado: {args.salida}')


if __name__ == '__main__':
    main()
