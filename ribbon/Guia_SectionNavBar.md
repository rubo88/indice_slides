# Section Nav Bar — Guía de Usuario
**Barra de Navegación por Secciones para PowerPoint**

**Autor:** Rubén Veiga Duarte ([ruben.veiga@bde.es](mailto:[ruben.veiga@bde.es])) - Mandame un correo si ves algún error o tienes alguna sugerencia.

**Descripción:** Genera automáticamente una barra de navegación visual en la parte superior de cada diapositiva. La barra muestra todas las secciones definidas como formas de chevron (flecha) enlazadas, permitiendo saltar directamente a cualquier sección durante la presentación.

**Última versión disponible:** 
[Link github](https://github.com/rubo88/indice_slides/blob/main/SectionNavBar.ppam) 
[Link a sharepoint](https://berso.sharepoint.com/sites/msteams_cc085a/_layouts/15/guestaccess.aspx?share=IgBgOPCuF1dISbJBVcR70cTSAdJ6x4JzUjjhIh3DBjilxfA&e=av4wcS)

**Fecha de última actualización:** 10/03/2026




---

## 1. ¿Qué hace esta macro?

Genera automáticamente una barra de navegación visual en la parte superior de cada diapositiva. La barra muestra todas las secciones definidas como formas de chevron (flecha) enlazadas, permitiendo saltar directamente a cualquier sección durante la presentación.

**Características principales:**

- Crea un chevron por cada sección definida en la presentación.
- La sección actual se resalta con fondo blanco; las demás quedan transparentes.
- Cada chevron es un hipervínculo que lleva a la primera diapositiva de esa sección.
- Re-ejecutar la macro actualiza la barra si cambian las secciones.

---

## 2. Secciones con comportamiento especial

La macro reconoce automáticamente tres tipos de secciones:

| Sección | Barra de navegación | Resaltado |
|---------|-------------------|-----------|
| **Portada** | No aparece | N/A — diapositiva limpia |
| **Apéndice / Apendice** | Sí aparece | Ninguna sección resaltada |
| **Cualquier otra** | Sí aparece | Su chevron se rellena de blanco |

Los nombres de sección se comparan sin distinguir mayúsculas/minúsculas y con o sin tilde.

---

## 4. Instalación como Add-in (.ppam)

El add-in se instala una sola vez y aparece automáticamente cada vez que abres PowerPoint.

### Pasos de instalación

1. Copia el archivo `.ppam` a `C/Trabajo`
2. Abre PowerPoint.
3. Ve a **Archivo → Opciones → Complementos**.
4. En la parte inferior, selecciona **"Complementos de PowerPoint"** en el desplegable y haz clic en **Ir...**
5. Haz clic en **Agregar nuevo...** y selecciona el archivo `.ppam`.
6. Marca la casilla junto al complemento y cierra.

> **Resultado:** Aparecerá una nueva pestaña **"Section Nav"** en la cinta de opciones con dos botones: **"Create Nav Bar"** y **"Remove Nav Bar"**.

---

## 5. Cómo usar

### Crear la barra de navegación

1. Abre tu presentación (asegúrate de que tiene secciones definidas).
2. Haz clic en la pestaña **"Section Nav"** en la cinta.
3. Pulsa el botón **"Create Nav Bar"**.
4. La macro procesará todas las diapositivas y mostrará un mensaje de confirmación.

### Actualizar la barra

Si cambias el nombre de una sección, añades nuevas secciones o reorganizas diapositivas, simplemente vuelve a ejecutar **"Create Nav Bar"**. La macro elimina las barras anteriores y las recrea con la estructura actual.

### Eliminar la barra

Pulsa el botón **"Remove Nav Bar"** para eliminar todas las formas de navegación de todas las diapositivas. Esto no afecta al resto del contenido.

---

## 6. Desinstalar

1. Ve a **Archivo → Opciones → Complementos**.
2. Selecciona **"Complementos de PowerPoint" → Ir...**
3. Desmarca o elimina el complemento de la lista.

La pestaña "Section Nav" desaparecerá inmediatamente del ribbon.
