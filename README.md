# Panel REM conectado a Google Sheets

Este proyecto reemplaza la carga manual de Excel por una lectura en vivo desde una pestaña de Google Sheets. Mantiene la regla del HTML original:

> Un registro se marca como **actividad no contabilizada en REM** cuando el concepto contiene `actividad no contabilizada en REM` o `actividad no contabilizada en RAM`.

El panel también identifica valores repetidos del ID y considera **“duplicado 100% sin REM”** a un ID repetido cuando *todas* sus filas cumplen la regla anterior.

## Importante: datos de salud y privacidad

La modalidad incluida es directa: el navegador descarga un CSV de Google Sheets. Por ese motivo la pestaña debe estar disponible para cualquier persona con el enlace como lectora. Úsela solo con datos **agregados, seudonimizados o sin identificadores personales**, y bajo las reglas de seguridad y gobernanza de su institución.

No publique una planilla con nombres, RUN, fichas clínicas, diagnósticos u otra información sensible. Para datos identificables, mantenga la planilla privada y use una alternativa con autenticación institucional (OAuth de Google Sheets API o un backend institucional). Esta versión no incluye credenciales ni guarda el contenido de la hoja: lo procesa solamente en el navegador que la abre.

## 1. Preparar Google Sheets

1. Cree una copia de trabajo anonimizada de la planilla y deje los títulos en la **fila 1**.
2. Para reconocer las variables, use preferentemente estos encabezados: `ID`, `Hora atención`, `Funcionario`, `Instrumento` y `Concepto`.
3. Copie el enlace de la pestaña. Se aceptan ambos formatos: `.../edit#gid=0` (pestaña compartida) y `.../pub?output=tsv` (pestaña publicada en TSV).
4. Solo si los datos no son sensibles, use el acceso por enlace o la publicación TSV. Ambos permiten la lectura sin autenticación; **no** use ninguno para datos identificables.

El panel detecta los encabezados por nombre, incluyendo `ATEN ID`, `HORA ATENCION`, `FUNCIONARIO`, `INSTRUMENTO` y `REM`. Si no los encuentra, conserva las posiciones del archivo anterior: **A** = ID, **E** = hora, **P** = funcionario, **Q** = instrumento y **AM** = concepto/detalle REM.

## 2. Probar el panel con Visual Studio Code

1. Descargue o copie esta carpeta a una ubicación de trabajo, por ejemplo `Documentos/GitHub/Actividadsinregistrorem`.
2. El panel ya tiene una fuente pública configurada en `config.public.js`; no hay que pegar enlaces para usarlo.
3. Si necesita probar otra fuente solo en su computador, copie `config.example.js` y renombre la copia como `config.local.js`. Luego complete esta línea:

   ```js
   globalThis.REM_SOURCE_URL = "https://docs.google.com/spreadsheets/d/e/.../pub?output=tsv";
   ```

4. Abra VS Code y seleccione **File / Archivo → Open Folder / Abrir carpeta**. Elija la carpeta del proyecto.
5. Instale la extensión **Live Server** desde Extensions si aún no la tiene.
6. Abra `index.html`; haga clic con el botón derecho y elija **Open with Live Server**. Se abrirá una dirección local, normalmente `http://127.0.0.1:5500`.
7. Los usuarios verán solo el botón **Actualizar datos**; no tendrán que pegar el enlace.

`config.public.js` permite que la página publicada funcione sin pedir enlaces al equipo. La fuente no se muestra en la interfaz, pero puede inspeccionarse en una web estática: no es un mecanismo de seguridad. `config.local.js` está en `.gitignore`, de modo que puede sobrescribir la fuente solo en su equipo sin subir esa modificación. Para datos sensibles, use OAuth/backend institucional.

No abra `index.html` con doble clic desde Finder: una dirección local servida por Live Server evita restricciones del navegador al consultar Google Sheets.

## 3. Qué mide y cómo leerlo

| Resultado | Definición operativa |
|---|---|
| Registros leídos | Filas de la pestaña, excluida la fila de encabezados. |
| ID únicos | Número de valores diferentes en la variable ID. |
| Filas duplicadas | Filas cuyo ID aparece más de una vez. No equivale necesariamente a un error clínico: depende de que el ID sea realmente único por atención. |
| Duplicados 100% sin REM | Filas pertenecientes a un ID repetido cuyas repeticiones completas están clasificadas como no contabilizadas en REM. |
| % no contabilizada | En el cruce Instrumento × REM: actividades no contabilizadas dividido por el total filtrado de ese instrumento. |

El cruce de variables y el gráfico responden a los filtros. La tabla conserva el número de fila original para facilitar la auditoría en Google Sheets. La exportación genera un Excel consolidado y una pestaña por funcionario.

Al escoger un **Instrumento**, el gráfico cambia automáticamente a **Funcionarios del instrumento** y el selector siguiente ofrece solo los funcionarios que tienen registros en ese instrumento. Por ejemplo, si escoge `Matrona`, verá el desglose de los funcionarios asociados a `Matrona` y no profesionales de otros instrumentos.

## 4. Cambiar reglas y nombres de columnas

Toda la configuración que se ajusta con más frecuencia está al inicio de [`app.js`](app.js):

```js
notCountedRemPhrases: [
  "actividad no contabilizada en rem",
  "actividad no contabilizada en ram"
]
```

Para sumar otra redacción usada por su establecimiento, agregue una tercera línea, por ejemplo:

```js
"actividad pendiente de contabilizar en rem"
```

Para reconocer un nuevo nombre de encabezado, agréguelo a `aliases`. Ejemplo para que **Instrumento** acepte también `Prestación`:

```js
instrument: {
  aliases: ["instrumento", "tipo instrumento", "nombre instrumento", "prestación"],
  fallbackColumn: "Q"
}
```

Otros cambios seguros:

* `previewLimit`: cantidad máxima de filas visibles en la tabla (por defecto 150).
* `chartLimit`: máximo de categorías mostradas en el gráfico (por defecto 8).
* `styles.css`: colores, tamaños y espaciados. Las variables principales están al inicio como `--teal`, `--coral` y `--gold`.
* `index.html`: textos visibles, títulos y opciones de filtro.

Después de guardar, Live Server actualizará la página. Si no ve el cambio, recargue con `Cmd + Shift + R` en macOS.

## 5. Trabajar con GitHub Desktop

Si ya tiene el repositorio `Actividadsinregistrorem` en GitHub:

1. Abra GitHub Desktop → **File → Clone repository** → pestaña **URL**.
2. Pegue `https://github.com/sergiosalazarcardenas-spec/Actividadsinregistrorem.git`, seleccione una carpeta local y pulse **Clone**.
3. En GitHub Desktop cree una rama con **Current branch → New branch**. Nombre sugerido: `feat/google-sheets-rem`.
4. Copie a la carpeta clonada los archivos de este proyecto: `index.html`, `styles.css`, `app.js`, `config.public.js`, `config.example.js`, `.gitignore` y este `README.md`. Sustituya el antiguo `index.html` solo después de conservar una copia de respaldo si la necesita. Mantenga `config.local.js` únicamente en su computador: no lo agregue al commit.
5. En GitHub Desktop seleccione **Repository → Open in Visual Studio Code** y siga la prueba del paso 2.
6. Vuelva a GitHub Desktop; en **Changes** revise el diff. Escriba como resumen: `feat: conectar panel REM a Google Sheets` y pulse **Commit to feat/google-sheets-rem**.
7. Pulse **Push origin**. Luego puede crear un Pull Request para revisar el cambio antes de incorporarlo a `main`.

GitHub Desktop permite clonar, crear ramas, confirmar y enviar cambios sin escribir comandos Git. La guía oficial explica el flujo de publicar/confirmar/enviar: [GitHub Desktop](https://docs.github.com/es/desktop/overview/creating-your-first-repository-using-github-desktop).

## 6. Publicar el panel

Antes de publicarlo en GitHub Pages, revise de nuevo que la planilla no contenga información personal o clínica. La URL de la hoja quedará visible para quien use el panel; no entregue en ella acceso a datos protegidos.

Para una planilla privada, el siguiente incremento técnico debe ser una implementación con inicio de sesión Google OAuth y el permiso de solo lectura `spreadsheets.readonly`, o un backend de la institución. No basta con ocultar un enlace o una clave API en `app.js`: ambos pueden ser vistos por quien carga una página web.

## Comprobación rápida antes de cada entrega

1. Pegue un enlace de una pestaña con datos de prueba anonimizados.
2. Compare el total de filas con `COUNTA`/conteo de Google Sheets.
3. Filtre “Actividad no contabilizada” y confirme una muestra manual contra el texto de `Concepto`.
4. Revise una fila “Duplicado 100% sin REM”: todas las filas con ese mismo ID deben estar clasificadas como no contabilizadas.
5. Exporte el resultado y confirme que el consolidado y las pestañas por funcionario contienen los filtros aplicados.
