# Foobar
Para la creación del archivo ejecutable seguir los siguientes pasos.

1. Clonar el repositorio.
2. Seguir las instrucciones del README
3. Instalar el paquete pyinstaller con le siguiente comando:

```bash
pip install pyinstaller
```
4. Ejecutar el siguiente comando en consola:

```bash
pyinstaller --noconfirm --onefile --windowed --icon "C:\Users\carlo\OneDrive\Documentos\SellOut ConAgro\Recursos\Icons\ConAgro.ico" --add-data "C:\Users\carlo\OneDrive\Documentos\SellOut ConAgro\Base clientes.xlsx;." --add-data "C:\Users\carlo\OneDrive\Documentos\SellOut ConAgro\ConAgro_icon_big.png;." --add-data "C:\Users\carlo\OneDrive\Documentos\SellOut ConAgro\ConAgro_icon_small.png;."  "C:\Users\carlo\OneDrive\Documentos\SellOut ConAgro\SellOut.py"
```
## Resultado.

Después del paso 4, tendremos como resultado dos carpetas, build y dist, en esta ultima encontraremos el archivo SellOut.exe