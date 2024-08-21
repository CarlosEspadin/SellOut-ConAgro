# SellOut and stock loading.
This repository contains the source code and executable file of the GUI project that cleans, transforms and uploads SellOut and Stock information from Syngenta distributors to the ConAgro web service.
## Instructions
The executable file is located in the following folder:
```bash
cd SellOut ConAgro\output\SellOut
```
To run the desktop application, the user must double-click the executable file or run the following command in the terminal:
```bash
.\output\SellOut\SellOut_Stocks.exe
```
## Input file layout SellOut

| Columns name | Type | Comment |
|----------|----------|----------|
| **folio**    | Varchar   | This field is required and is the unique identifier for each record.   |
| **fecha_Facturacion**   | Date   | This field is required and its format is as follows: YYYY-MM-DD, it cannot contain null records.  |
| **volumen_Facturado**   | Fload   | This field is required.   |
| **unidad_Medida**   | Varchar   | This field is optional.   |
| **valorT_Facturado**   | Fload   | This field is required.   |
| **rfc**   | Varchar   | This field is required.   |
| **clave_Distribuidor**   |    Int  | This field is required.   |
| **nombre_distribuidor**   | Varchar  | This field is required.   |
| **nombre_Cliente**   | Varchar   | This field is optional.   |
| **codeProduct_Distribuidor**   | Varchar   | This field is required.   |
| **localidad**   | Varchar   | This field is optional.  |
| **sucursal**   | Varchar   | This field is optional.  |
| **nombre_Vendedor_Distribuidor**   | Varchar   | This field is optional.   |
| **linea_Producto**   | Varchar   | This field is optional.   |
| **marca**   | Varchar   | This field is required. Your value should always be Syngenta**   |
| **pais**   | Varchar   | This field is optional.   |
| **numero_Convenio**   | Varchar  | This field is optional.    |

## Input file layout Stock

| Columns name | Type | Comment |
|----------|----------|----------|
| **fecha_Inventario**   | Date   | This field is required and its format is as follows: YYYY-MM-DD, it cannot contain null records.  |
| **linea_Negocio**   | Varchar   | This field is optional.   |
| **codeProduct_Distribuidor**   | Varchar   | This field is optional.   |
| **presentacion**   | Fload   | This field is required.   |
| **unidad_Medida**   | Varchar   | This field is optional.   |
| **volumen_Inventario**   | Fload  | This field is required.   |
| **clave_Distribuidor**   | Int  | This field is required.   |
| **nombre_Distribuidor**   | Varchar   | This field is required.  |
| **almacen**   | Varchar   | This field is optional.  |
| **municipio**  | Varchar   | This field is optional.  |
| **no_ShipTo**   | Varchar   | This field is optional.  |
| **pais**   | Varchar   | This field is required.   |
| **inventario_EnTransito**   | Fload   | This field is optional.   |


## Warning
To run the “SellOut_Stocks.exe” application it is necessary not to delete the “_internal” folder and the “Base Clientes.xlsx” file, otherwise an error will occur.