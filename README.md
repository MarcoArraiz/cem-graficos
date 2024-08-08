
# Graficador

## Descripción del Proyecto

Este script está diseñado para generar gráficos de barras en Excel a partir de datos de mineralogía y geoquímica en el rubro de la minería. Utiliza los datos proporcionados en un archivo Excel, organiza la información en un DataFrame y crea gráficos personalizados para cada mineral o elemento químico, asignando colores específicos a cada uno.

## Funcionalidades

- **Generación de Gráficos de Barras:** Crea gráficos de barras para cada mineral o elemento químico con colores específicos.
- **Inclusión de Todos los Metros:** Asegura que todos los metros desde el mínimo hasta el máximo estén incluidos en el gráfico, rellenando con ceros donde falten datos.
- **Filtrado de Valores Cero:** Permite la opción de excluir gráficos para columnas cuyos valores son todos cero.
- **Unidades de Medida:** Convierte los valores a PPM si todos los valores son menores a 0.1%.

## Estructura del Proyecto

### Archivos y Directorios

- `main.py`: Archivo principal del script.
- `Entregable Proyecto.xlsx`: Archivo Excel con los datos de entrada.
- `Grafico_Barra_Geoquimica.xlsx`: Archivo de salida con los gráficos de barras para los datos geoquímicos.
- `Grafico_Barra_Mineralogia.xlsx`: Archivo de salida con los gráficos de barras para los datos de mineralogía.
- `README.md`: Archivo de documentación del proyecto.

### Requisitos

- Python 3.x
- Pandas
- XlsxWriter

### Instalación

1. **Clonar el repositorio:**
   ```bash
   git clone https://github.com/MarcoArraiz/cem-graficos
   ```
2. **Navegar al directorio del proyecto:**
   ```bash
   cd cem-graficos
   ```
3. **Instalar las dependencias:**
   ```bash
   pip install -r requirements.txt
   ```

### Uso

1. **Preparar el archivo de datos:**
   Asegúrate de tener el archivo `Entregable Proyecto.xlsx` en el directorio del proyecto. Este archivo debe contener dos hojas: `Geoquímica` y `Mineralogía`.

2. **Ejecutar el script:**
   ```bash
   python main.py
   ```

3. **Revisar los archivos de salida:**
   Los gráficos de barras generados se guardarán en dos archivos: `Grafico_Barra_Geoquimica.xlsx` y `Grafico_Barra_Mineralogia.xlsx`.

### Ejemplo de Estructura de Carpetas

```
project-root/
├── Entregable Proyecto.xlsx
├── Grafico_Barra_Geoquimica.xlsx
├── Grafico_Barra_Mineralogia.xlsx
├── main.py
└── README.md
```

### Detalles del Código

El script principal `main.py` incluye la función `generar_grafico_barra`, que se encarga de crear gráficos de barras en Excel. A continuación se describen las partes principales del script:

- **Definición de Colores:** Se define un diccionario `element_colors` que asigna colores específicos a cada mineral o elemento.
- **Función `generar_grafico_barra`:** Esta función toma un DataFrame, el nombre del archivo de salida y una opción para filtrar columnas con valores cero.
  - Se asegura de incluir todos los metros desde el mínimo hasta el máximo, rellenando con ceros donde sea necesario.
  - Excluye las columnas "TOTAL", "Litología" y "Zona Mineral".
  - Crea una hoja de resumen con todos los gráficos generados.

- **Ejecución del Script:**
  - Carga los datos desde el archivo Excel `Entregable Proyecto.xlsx`.
  - Genera gráficos para la hoja `Geoquímica`, excluyendo elementos con todos sus valores en cero.
  - Genera gráficos para la hoja `Mineralogía` sin filtrar valores cero.

### Contacto

Para cualquier consulta o soporte, puedes contactarme a través de:

- Email: marcoarraiz@gmail.com
- GitHub: [MarcoArraiz](https://github.com/MarcoArraiz)

## Licencia

Este proyecto está licenciado bajo la Licencia MIT. Ver el archivo `LICENSE` para más detalles.

---

Este README proporciona una guía completa sobre cómo utilizar el script para generar gráficos de barras a partir de datos de mineralogía y geoquímica. Si tienes alguna duda o necesitas asistencia, no dudes en contactarme.
