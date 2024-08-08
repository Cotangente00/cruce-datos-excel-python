- Nota Importante: la aplicación está desarrollada única y exclusivamente para la plantilla **INFORME SOLICITUDES,** por lo tanto, si esta plantilla es modificada, **el código de la aplicación también debe ser adaptado para que no haya incongruencias con los datos.**



La aplicación únicamente se puede ejecutar mediante un ambiente CMD, por limitaciones de los equipos de la empresa.
La aplicación está pensada para ser incluida al SIC su uso será destinado exclusivamente a todas las gestoras de servicios pertenecientes al área de operaciones.

Pasos para hacer un ejecutable CMD:

1. Abrir un bloc de notas.
2. Copiar el contenido del archivo "ejecutar_programa.cmd" y pegarlo en el bloc de notas.
3. Ajustar la ruta a la ubicación del directorio del programa.
4. Guardar el archivo como formato CMD (IMPORTANTE).
5. El archivo CMD tiene que estar ubicado en la ruta local de la aplicación.
6. Click derecho en el archivo CMD para crear un acceso directo y ubicar este acceso directo a una ruta a convenir para un acceso eficaz
7. Consideraciones: Tener Python instalado, tener todas las dependencias y librerías necesarias instaladas

No hay entorno Virtual por las limitaciones que tienen configuradas los equipos de la empresa.

Comando para instalar todas las dependencias y librerías necesarias: pip install -r requirements.txt

De lo contrario, instalar las dependencias y librerías una a una, estas librerías son: openpyxl, et-xmlfile y xlrd. Comando para instalar programa de uno uno: "pip install `<nombre de la librería>`"

Comando para convertir correctamente el proyecto a un ejecutable .exe: pyinstaller --onefile --windowed --name="Informe de solicitudes y disponibilidad de expertas" --add-data "icon.ico;." programa.py

Para poder ejecutar este comando, es necesario tener instalada la librería de Pyinstaller con el comando "pip install pyinstaller"
