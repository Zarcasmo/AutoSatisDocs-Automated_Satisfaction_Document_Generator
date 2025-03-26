
# **AutoSatisDocs-Automated Satisfaction Document Generator**

ActaGen es una herramienta automatizada para generar actas de satisfacción en formatos Word y PDF. Utiliza una base de datos en Excel para alimentar la información necesaria y firmas digitalizadas almacenadas en una carpeta dentro del repositorio.

## Características

- Generación automática de actas de satisfacción en formatos Word y PDF.
- Integración con bases de datos en Excel para la entrada de datos.
- Inclusión de firmas digitalizadas desde una carpeta específica.

## Requisitos

- Python 3.x
- Librerías: `pandas`, `openpyxl`, `python-docx`, `fpdf`

## Instalación

1. Clona este repositorio:
    ```bash
    git clone https://github.com/tu-usuario/ActaGen.git
    ```
2. Navega al directorio del proyecto:
    ```bash
    cd ActaGen
    ```
3. Instala las dependencias:
    ```bash
    pip install -r requirements.txt
    ```

## Uso

1. Coloca tu archivo de base de datos en Excel en la carpeta `data`.
2. Coloca las firmas digitalizadas en la carpeta `signatures`.
3. Ejecuta el script principal:
    ```bash
    python generate_actas.py
    ```

## Contribuciones

¡Las contribuciones son bienvenidas! Por favor, abre un issue o envía un pull request para discutir cualquier cambio que te gustaría realizar.

## Licencia

Este proyecto está licenciado bajo la Licencia MIT. Consulta el archivo `LICENSE` para más detalles.
 
