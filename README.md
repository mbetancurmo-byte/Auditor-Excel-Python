# Auditor Excel: conciliación de distribución y contabilidad

Proyecto de portafolio en Python para comparar dos Excel por identificación y concepto. Los ejemplos incluidos son ficticios: no son información de ninguna entidad ni conclusiones de auditoría.

## Inicio en Windows (PowerShell)

Instala Python 3.10 o superior. Extrae el ZIP y abre la carpeta auditor_excel en VS Code. En su terminal ejecuta:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe conciliar.py ejemplos/distribucion.xlsx ejemplos/contabilidad.xlsx --salida mi_informe.xlsx
```

No necesitas activar el entorno virtual. La primera instalación de dependencias requiere acceso a Internet. El análisis se ejecuta localmente.

## Formato de tus datos

En la primera fila deben estar las columnas `identificacion`, `concepto` y `valor`. Puede haber otras columnas; no intervienen en el cruce. Se usa la primera hoja salvo que indiques otra:

```powershell
.\.venv\Scripts\python.exe conciliar.py distribucion.xlsx contabilidad.xlsx --hoja-distribucion Datos --hoja-contabilidad Datos --salida revision.xlsx --tolerancia 0.01
```

Guarda las identificaciones como texto para conservar ceros iniciales. Los valores deben ser números o texto como `1234.56`, sin símbolos ni separadores de miles y con máximo dos decimales. Se aceptan negativos. No se aceptan fórmulas en las columnas requeridas: prepara una copia con valores comprobados.

Antes de comparar, verifica que ambos archivos correspondan a la misma moneda, periodo y alcance. El programa no puede determinar esa equivalencia.

## Qué hace el código

1. `leer`: comprueba columnas, conserva números de fila y separa registros inválidos.
2. `importe`: convierte los importes a centavos enteros para sumar y comparar con exactitud.
3. `agrupar`: suma por identificación y concepto. Recorta espacios externos; mantiene mayúsculas, acentos y demás caracteres.
4. `cruzar`: compara grupos y marca coincidencias, diferencias y ausencias. Diferencia = distribución menos contabilidad.
5. `ejecutar`: crea el informe, registra huellas SHA-256 de las entradas y conserva las bases válidas.
6. `main`: permite ejecutar el programa desde la terminal y elegir hojas, salida y tolerancia.

## Informe generado

| Hoja | Contenido |
| --- | --- |
| Resumen | Totales válidos, exclusiones y cantidad de grupos por estado |
| Cruce | Valores comparados y filas de origen |
| Invalidos | Filas excluidas, archivo, hoja y motivo |
| Posibles_duplicados | Repeticiones de identificación, concepto y valor |
| Base_distribucion | Registros válidos de distribución |
| Base_contabilidad | Registros válidos de contabilidad |
| Trazabilidad | Fecha, reglas y huellas de los archivos |

Las repeticiones no se eliminan: pueden ser operaciones legítimas. Si existen filas inválidas, el proceso se marca PARCIAL y los totales no representan la totalidad de las entradas. Una coincidencia por grupo puede ocultar diferencias entre operaciones; para comprobarlas necesitas una clave de transacción común. El programa no analiza fechas ni acredita pagos, pérdidas o fraude. Las huellas identifican los archivos utilizados; no certifican su autenticidad.

Los cálculos internos usan centavos enteros. Excel recibe importes numéricos para revisión; utiliza cantidades dentro de la precisión habitual de Excel (15 dígitos significativos). El programa exige una salida nueva para evitar sobrescrituras.

## Demostración comprobable

Los ejemplos contienen: 310 de distribución válida, 280 de contabilidad y diferencia neta 30. Hay dos grupos coincidentes, uno con diferencia, uno exclusivo de cada archivo, una fila inválida y un grupo de posibles duplicados. Consulta `ejemplos/informe_demo.xlsx`.

Para ejecutar las pruebas:

```powershell
.\.venv\Scripts\python.exe -m unittest discover -s tests -v
```

Para recrear los ejemplos en una carpeta sin los archivos de entrada, ejecuta `python generar_ejemplos.py`.

## Cómo mostrarlo en tu portafolio

Puedes crear un repositorio llamado `auditor-excel-python` y publicar este código, README, pruebas y ejemplos ficticios. No publiques información real de auditoría. Explica en una entrevista: «Desarrollé una herramienta que compara dos fuentes financieras, detecta diferencias y conserva trazabilidad hasta las filas originales; validé el comportamiento con pruebas automatizadas».

Demuestra lectura de Excel, validación de datos, estructuras de Python, tratamiento monetario, automatización, trazabilidad y pruebas. Es una base de portafolio; la contratación depende también de tu capacidad de explicarlo y adaptarlo.
