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
pyinstaller --noconfirm --onedir --windowed --icon "C:\Users\carlo\OneDrive\Documentos\SellOut ConAgro\ConAgro.ico" --add-data "C:\Users\carlo\OneDrive\Documentos\SellOut ConAgro\agriculture_6739552.png;." --add-data "C:\Users\carlo\OneDrive\Documentos\SellOut ConAgro\ConAgro.ico;." --add-data "C:\Users\carlo\OneDrive\Documentos\SellOut ConAgro\ConAgro_icon_big.png;." --add-data "C:\Users\carlo\OneDrive\Documentos\SellOut ConAgro\ConAgro_icon_small.png;." --add-data "C:\Users\carlo\OneDrive\Documentos\SellOut ConAgro\home_2115185.png;." --add-data "C:\Users\carlo\OneDrive\Documentos\SellOut ConAgro\menu.png;." --add-data "C:\Users\carlo\OneDrive\Documentos\SellOut ConAgro\onboarding_14753055.png;." --add-data "C:\Users\carlo\OneDrive\Documentos\SellOut ConAgro\README.md;." --add-data "C:\Users\carlo\OneDrive\Documentos\SellOut ConAgro\review-document_14752610.png;." --add-data "C:\Users\carlo\OneDrive\Documentos\SellOut ConAgro\setting_1146744.png;." --add-data "C:\Users\carlo\OneDrive\Documentos\SellOut ConAgro\workflow_14254834.png;."  "C:\Users\carlo\OneDrive\Documentos\SellOut ConAgro\SellOut.py"
```
## Resultado.

Después del paso 4, tendremos como resultado dos carpetas, build y dist, en esta ultima encontraremos el archivo SellOut.exe