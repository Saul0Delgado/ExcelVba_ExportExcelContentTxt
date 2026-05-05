# ExportExcelToTxt

AddIn para exportar el contenido de los libros de Excel a un archivo de texto.

Útil para herramientas de comparación de versiones que no entienden Excel de forma nativa.

Genera un archivo de texto en el mismo folder, con el mismo nombre que el archivo a exportar, con extensión `.TXTXL.txt`


# Instalación.

1. Bajar el [último release](http://15.35.10.167:9999/El_Servas/ExportExcelToTxt/tags)
2. Copiar el archivo .XLA en la ruta Autostart de Excel: `%appdata%/Microsoft/Excel/XLSTART`
3. Reiniciar Excel.


# Uso

En un archivo abierto en Excel, presionar Ctrl-T para iniciar el proceso de exportación.


# Generación manual .XLA

Si se desea generar el archivo .XLA a partir del fuente, ejecutar `\build\build.bat`.
  
El archivo .XLA se genera en el mismo folder que el fuente (.XLS)
