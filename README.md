# Auditoría BIM: instrumento de investigación

Cuestionario web de una sola pantalla sobre cómo los coordinadores BIM auditan sus modelos de
Revit antes de una reunión de coordinación o de una entrega.

**Responder aquí:** https://matbimdev.github.io/bim-audit-survey/

Es el instrumento de evidencia directa del Hito 1 del Trabajo General del Diplomado de IA
en Ingeniería y Construcción. Alimenta las secciones de usuario y de evidencia del informe
de planteamiento del proyecto BIM Agent.

Tiempo de respuesta: unos 10 minutos. Bilingüe, español e inglés, con selector arriba.

## Privacidad

Los datos personales no se publican. Las respuestas se citan de forma anónima en el trabajo
académico, indicando solo el rol y el tipo de empresa. El correo se pide únicamente por si hace
falta una repregunta. Hay una casilla de consentimiento obligatoria antes de poder enviar.

## Cómo está hecho

Un solo archivo, `index.html`, sin dependencias ni build. GitHub Pages lo sirve tal cual.

- El formulario se construye desde un esquema declarativo dentro del propio script, así que cada
  pregunta y cada opción viven en un solo sitio.
- Las opciones de la pregunta principal se barajan por encuestado y ese orden queda fijo durante
  toda la sesión. Lo seleccionado se guarda por identificador estable, nunca por posición, y el
  orden que le tocó a cada persona se registra junto a la respuesta.
- El selector de idioma solo reemplaza texto. Los valores escritos no se tocan, así que cambiar de
  idioma a mitad del formulario no borra nada.
- Al abrir, la página consulta el endpoint. Si no responde, bloquea el formulario con un aviso en
  vez de dejar escribir y perder la respuesta al final.
- Si el envío falla, la respuesta no se pierde: aparece un panel con reintentar, descargar en JSON
  y enviar por correo.

## Puesta en marcha del almacenamiento

Las respuestas van a una hoja de cálculo de Google a través de un Apps Script. Pasos:

1. Crear una hoja de cálculo nueva en Google Drive.
2. Dentro de la hoja, menú **Extensiones → Apps Script**.
3. Borrar el contenido de `Código.gs` y pegar el script de más abajo. Guardar.
4. **Implementar → Nueva implementación**, tipo **Aplicación web**.
5. Ejecutar como **Yo**. Quién tiene acceso: **Cualquier usuario**. Implementar y autorizar.
6. Copiar la URL que termina en `/exec`.
7. Pegarla en `index.html`, en `CONFIG.endpoint`, al inicio del bloque de script.

Al cambiar el script después hay que ir a **Implementar → Gestionar implementaciones**, editar la
implementación y publicar una **versión nueva**. Si no, el cambio no llega a la URL `/exec`.

### Por qué el POST va como texto plano

Un `POST` con `Content-Type: application/json` dispara una petición previa `OPTIONS` que Apps
Script no atiende, y falla siempre. La página envía `text/plain;charset=utf-8`, que es una
petición simple sin verificación previa. `/exec` redirige a `script.googleusercontent.com` y esa
respuesta final sí trae las cabeceras que permiten leer el JSON de confirmación.

### Sobre exponer el endpoint

La URL del script queda a la vista en este repositorio, así que cualquiera podría enviarle datos.
El script descarta lo que no traiga el identificador de formulario correcto y lo que rellene el
campo trampa oculto, que es lo que suelen hacer los bots. Para un instrumento de 3 a 10 respuestas
es suficiente. Si aparece basura, basta con publicar una implementación nueva: la URL anterior
deja de servir.

### El script

```javascript
/**
 * Backend de la encuesta "Auditoría BIM: instrumento de investigación".
 * Va dentro de una hoja de cálculo: Extensiones -> Apps Script.
 * Implementación: Ejecutar como "Yo", acceso "Cualquier usuario".
 */

var FORM_ID = 'bim-audit-survey-v1';
var HOJA = 'respuestas';

var COLUMNAS = [
  'enviado_en', 'idioma', 'consentimiento',
  'nombre', 'correo', 'puesto', 'empresa', 'pais',
  'revit_versiones', 'herramientas_actuales',
  'que_revisas', 'tiempo_valor', 'tiempo_unidad', 'frecuencia',
  'reglas_escritas', 'reglas_cambian', 'que_se_escapa',
  'capacidades_deseadas', 'capacidades_otra_texto', 'formato_resultado', 'permiso_modificar',
  'preocupaciones', 'instalacion',
  'tiempo_liberado', 'experiencia_ia',
  'orden_capacidades', 'json_crudo'
];

function doGet() {
  return json_({ ok: true, status: 'ready' });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    var p = JSON.parse(e.postData.contents);
    if (p.form_id !== FORM_ID) return json_({ ok: false, error: 'form_id' });
    if (p.trampa) return json_({ ok: true, ignorado: true });

    var hoja = hoja_();
    var r = p.respuestas || {};
    var t = r.tiempo_revision || {};

    hoja.appendRow([
      p.enviado_en || new Date().toISOString(),
      p.idioma || '',
      p.consentimiento ? 'si' : 'no',
      r.nombre || '', r.correo || '', r.puesto || '', r.empresa || '', r.pais || '',
      unir_(r.revit_versiones), unir_(r.herramientas_actuales),
      r.que_revisas || '',
      (t.valor === null || t.valor === undefined) ? '' : t.valor,
      t.unidad || '',
      r.frecuencia || '', r.reglas_escritas || '', r.reglas_cambian || '', r.que_se_escapa || '',
      unir_(r.capacidades_deseadas), r.capacidades_deseadas_otra_texto || '',
      unir_(r.formato_resultado), r.permiso_modificar || '',
      unir_(r.preocupaciones), r.instalacion || '',
      r.tiempo_liberado || '', r.experiencia_ia || '',
      unir_(p.orden_capacidades),
      JSON.stringify(p)
    ]);

    return json_({ ok: true, row: hoja.getLastRow() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignored) {}
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function unir_(v) {
  return (v && v.join) ? v.join(' | ') : (v || '');
}

function hoja_() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = libro.getSheetByName(HOJA) || libro.insertSheet(HOJA);
  if (hoja.getLastRow() === 0) {
    hoja.appendRow(COLUMNAS);
    hoja.setFrozenRows(1);
  }
  return hoja;
}
```

## Cómo llega cada respuesta a la hoja

Una fila por envío. Las preguntas de selección múltiple se guardan como identificadores separados
por ` | `. Las de opción única, como un identificador suelto. El tiempo de revisión se parte en
dos columnas, cantidad y unidad. La última columna guarda el JSON completo por si hiciera falta
recuperar algo que el mapeo de columnas no contemple.

Identificadores de la pregunta principal, en el orden a) a m) del instrumento: `warnings`,
`parametros_vacios`, `familias`, `nombrado`, `metrados`, `parametros_elemento`, `busqueda`,
`planos_vistas`, `vinculos_coordenadas`, `interferencias`, `radiografia`, `info_proyecto`,
`exportar`, más `otra`.

## Desarrollo local

```bash
py -m http.server 8080 --bind 127.0.0.1
```

Con `CONFIG.endpoint` vacío la página se abre bloqueada, que es el comportamiento correcto cuando
no hay dónde guardar. Para probar el flujo completo hay que apuntarlo a un endpoint real o a un
simulador local que responda `{"ok": true}` a `GET` y a `POST`.
