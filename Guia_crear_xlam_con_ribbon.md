# Cómo crear un Add-in de Excel (.xlam) con pestaña en el Ribbon

**Partiendo de un archivo .bas con tus macros**

---

## Paso 1: Crear el archivo .xlam base

1. Abre Excel → **Libro en blanco**.
2. Pulsa **Alt+F11** para abrir el Editor de VBA.
3. En el panel izquierdo, haz clic derecho sobre **VBAProject (Libro1)** → **Importar archivo...** → selecciona tu archivo `.bas`.
4. Si tienes varios `.bas`, repite para cada uno.
5. Cierra el Editor de VBA.
6. Ve a **Archivo → Guardar como** → en tipo selecciona **Complemento de Excel (.xlam)**.
7. Ponle un nombre (por ejemplo `MisMacros.xlam`) y guárdalo donde quieras.
8. **Cierra Excel completamente.**

> En este punto tienes un .xlam funcional, pero sin pestaña en el ribbon. Las macros solo serían accesibles desde Alt+F8.

---

## Paso 2: Crear los callbacks para el ribbon

El ribbon necesita que cada botón llame a un `Sub` que acepte un parámetro `IRibbonControl`. Tus macros originales no lo tienen, así que hay que crear un módulo puente.

Por cada macro que quieras poner en el ribbon, crea un wrapper así:

```vba
Public Sub RibbonMiMacro(control As IRibbonControl)
    MiMacro    ' ← llama a tu macro original
End Sub
```

**Ejemplo completo** — si tienes estas macros:

```
Sub HacerCosaA()
Sub HacerCosaB()
Sub HacerCosaC()
```

Creas un archivo `RibbonCallbacks.bas` con:

```vba
Attribute VB_Name = "RibbonCallbacks"
Option Explicit

Public Sub RibbonHacerCosaA(control As IRibbonControl)
    HacerCosaA
End Sub

Public Sub RibbonHacerCosaB(control As IRibbonControl)
    HacerCosaB
End Sub

Public Sub RibbonHacerCosaC(control As IRibbonControl)
    HacerCosaC
End Sub
```

Importa este `.bas` en tu `.xlam` igual que en el Paso 1 (abrir el .xlam, Alt+F11, importar, guardar, cerrar Excel).

---

## Paso 3: Crear el archivo XML del ribbon

Crea un archivo de texto llamado `customUI14.xml` con la definición de tu pestaña. Esta es la estructura:

```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="miTab" label="Mi Pestaña">
        <group id="grp1" label="Grupo 1">
          <button id="btn1"
                  label="Cosa A"
                  onAction="RibbonHacerCosaA"
                  imageMso="HappyFace"
                  size="large"
                  screentip="Título del tooltip"
                  supertip="Descripción larga del tooltip." />
          <button id="btn2"
                  label="Cosa B"
                  onAction="RibbonHacerCosaB"
                  imageMso="Refresh"
                  size="large" />
        </group>
        <group id="grp2" label="Grupo 2">
          <button id="btn3"
                  label="Cosa C"
                  onAction="RibbonHacerCosaC"
                  imageMso="Copy"
                  size="normal" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

**Reglas importantes del XML:**

- Cada `id` debe ser único en todo el archivo.
- El valor de `onAction` debe coincidir **exactamente** con el nombre del `Sub` en VBA (el wrapper, no la macro original).
- `size="large"` muestra icono grande con texto debajo; `size="normal"` muestra icono pequeño con texto al lado.
- `imageMso` es el nombre del icono de Office. Busca iconos disponibles en: [bert-toolkit.com/imagemso-list.html](https://bert-toolkit.com/imagemso-list.html)
- `screentip` y `supertip` son opcionales (tooltip corto y largo).

---

## Paso 4: Inyectar el XML con Office RibbonX Editor

1. **Descarga** Office RibbonX Editor desde: [github.com/fernandreu/office-ribbonx-editor/releases](https://github.com/fernandreu/office-ribbonx-editor/releases) (busca el `.exe` o `.zip` de la última release).

2. **Asegúrate de que Excel está completamente cerrado** (el .xlam no puede estar abierto en Excel mientras lo editas).

3. Abre **Office RibbonX Editor**.

4. Haz clic en **File → Open** y selecciona tu archivo `.xlam`.

5. En el panel izquierdo aparecerá tu archivo. Haz **clic derecho** sobre él → **Insert Office 2010+ Custom UI Part**.

   > Aparecerá un nodo `customUI14.xml` debajo del archivo.

6. Haz clic en ese nodo `customUI14.xml`. En el panel derecho (editor de texto) **pega el contenido completo** de tu archivo XML.

7. Haz clic en el botón **Validate** (icono de check verde, arriba) para comprobar que el XML no tiene errores.

8. Si la validación es correcta, haz clic en **File → Save** (o Ctrl+S).

9. Cierra Office RibbonX Editor.

---

## Paso 5: Probar el resultado

1. Abre Excel.
2. Ve a **Archivo → Opciones → Complementos**.
3. En la parte inferior: **Administrar → Complementos de Excel → Ir...**
4. Si tu .xlam no aparece en la lista, haz clic en **Examinar...** y búscalo.
5. Marca la casilla y cierra.
6. Tu nueva pestaña debería aparecer en el ribbon.

---

## Resumen visual del proceso

```
  .bas (tus macros)
       │
       ▼
  Excel: importar .bas → Guardar como .xlam → Cerrar Excel
       │
       ▼
  .bas (callbacks con IRibbonControl) → Importar en el .xlam → Guardar → Cerrar Excel
       │
       ▼
  customUI14.xml (definición del ribbon)
       │
       ▼
  Office RibbonX Editor: abrir .xlam → Insert Custom UI Part → pegar XML → Validate → Save
       │
       ▼
  Excel: Archivo → Opciones → Complementos → Activar el .xlam
       │
       ▼
  ✓ Tu pestaña aparece en el ribbon
```

---

## Distribuir a compañeros

Simplemente envíales el archivo `.xlam` terminado. El XML del ribbon ya está incrustado dentro, así que ellos solo tienen que hacer el Paso 5 (activar el complemento). No necesitan Office RibbonX Editor ni hacer nada más.

Si quieres que no les salte ningún aviso de seguridad, diles que añadan la carpeta donde guardan el `.xlam` como **ubicación de confianza**: Archivo → Opciones → Centro de confianza → Configuración → Ubicaciones de confianza → Agregar nueva ubicación.
