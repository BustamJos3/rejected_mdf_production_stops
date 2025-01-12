# Archivos no encontrados en script - Usar la Ruta Absoluta del Script

Si el script encuentra los archivos en el notebook pero no en el terminal, el problema probablemente esté relacionado con la **ruta de trabajo** (directorio actual) desde la cual se está ejecutando el script en el terminal. Cuando se ejecuta un notebook, la ruta de trabajo suele estar configurada en el directorio del propio notebook, mientras que al ejecutar un script desde el terminal, la ruta de trabajo podría diferir.

Para asegurarte de que el script utiliza siempre la ruta correcta para `data_plots`, intenta lo siguiente:

### Solución 1: Usar la Ruta Absoluta del Script
Modifica el script para que obtenga la ruta de `data_plots` en función de la ubicación real del archivo del script, en lugar de la ruta de trabajo actual.

1. Usa `Path(__file__).parent` para obtener la ubicación del script en el sistema de archivos.
2. Configura `data_plots` como una carpeta dentro de esa ubicación.

Modifica el código de la siguiente manera:

```python
from pathlib import Path
import pandas as pd

# Definir la ruta de 'data_plots' en base a la ubicación del script
script_dir = Path(__file__).parent  # Obtiene el directorio del script
directory = script_dir / 'data_plots'  # Define la ruta de data_plots relativa al script
directory.mkdir(exist_ok=True)  # Crear si no existe
print("Ruta absoluta de data_plots:", directory)

# Buscar archivos .xlsx que contienen 'obj'
matching_files = list(directory.glob("*obj*.xlsx"))
print("Archivos encontrados:", matching_files)

# Procesar archivos encontrados
dict_data_pointer = {}
for file_path in matching_files:
    df = pd.read_excel(file_path)
    df_name = file_path.stem.split("obj_")[1]
    dict_data_pointer[df_name] = df

# Verificar las claves
print("Claves en dict_data_pointer:", list(dict_data_pointer.keys()))
