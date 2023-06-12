# Proyecto de Comprobación de Combinaciones con PrestaShop
https://wakatime.com/badge/user/23a2ba12-aa90-4e65-bc81-adc19381561f/project/493ccce3-8370-4a48-9d28-0462c295c9a8.svg
Este proyecto tiene como objetivo principal comprobar las combinaciones de productos utilizando un archivo de Excel como fuente de datos. Se compararán los códigos EAN13 en el archivo de Excel con la API de PrestaShop para determinar si se deben subir o no. Además, se proporcionará una interfaz gráfica para facilitar el proceso.

![Captura de pantalla 2023-06-12 124935](https://github.com/Juanies/Comprobador-combinaciones-duplicadas-Excel---Prestashop/assets/80675013/bd0a2c93-59d4-4531-8b38-4e2d0502c995)

## Requisitos previos

Antes de ejecutar este proyecto, asegúrate de tener instalado lo siguiente:

- Python (versión 3.10.11)
- Bibliotecas adicionales (se especificarán en el archivo `requirements.txt`)

## Instalación

1. Clona este repositorio en tu máquina local.
2. Navega hasta la carpeta del proyecto.
3. Instala las dependencias utilizando el siguiente comando:

   ```shell
   pip install -r requirements.txt

## USO

Ejecuta el siguiente comando para iniciar la interfaz gráfica:
```shell
python index.py
````
Selecciona la tienda en la que deseas comprobar las combinaciones.

Haz clic en el botón "Analizar" para comenzar el proceso de comparación con la API de PrestaShop.

El programa mostrará los resultados y proporcionará la opción de eliminar duplicados en el archivo de Excel.
```diff
- [Recuerda cambiar las columnas y los datos del JSON]
```
## Contribución

Si deseas contribuir a este proyecto, sigue los pasos a continuación:

1. Haz un fork de este repositorio.
2. Crea una rama con la nueva funcionalidad: git checkout -b nueva-funcionalidad.
3. Realiza los cambios necesarios y guarda los archivos.
4. Realiza un commit de tus cambios: git commit -m "Agrega nueva funcionalidad".
5. Envía tus cambios al repositorio remoto: git push origin nueva-funcionalidad.
6. Crea una Pull Request en GitHub.

## Créditos
Este proyecto ha sido desarrollado por [Juanies#6389](https://github.com/Juanies)

![image](https://github.com/Juanies/Comprobador-combinaciones-duplicadas-Excel---Prestashop/assets/80675013/c895ad3f-53ff-4632-8e2e-2ebcb617a99c)


