# Automatización en Google Sheets con Google Apps Script — asignación y trazabilidad

> ## ⚠ DATOS 100 % FICTICIOS
> El libro de ejemplo [`datos-ejemplo/SOLICITUDES_FICTICIAS.xlsx`](datos-ejemplo/SOLICITUDES_FICTICIAS.xlsx)
> fue **generado artificialmente** con [`herramientas/generar_libro_ficticio.py`](herramientas/generar_libro_ficticio.py).
> Las razones sociales (`EMPRESA FICTICIA 0001 S.A.`), los RUC (prefijo `99`, provincia inexistente), los
> documentos (`DOC-FICT-...`) y los funcionarios (`FUNCIONARIO A–D`) **no corresponden a personas ni
> organizaciones reales**. El repositorio no incluye datos de ninguna institución.

## Qué es
Una implementación de **automatización operativa en Google Sheets** con **Google Apps Script**, orientada a
la gestión de solicitudes: registra eventos, sincroniza contexto entre hojas y sugiere responsables a partir
de reglas centralizadas. Convierte una hoja de cálculo tradicional en un **sistema reactivo** que ejecuta
lógica de negocio en tiempo real ante las acciones del usuario.

## Qué demuestra
- **Automatización por eventos** (`onEdit`) sobre una hoja de uso diario.
- **Configuración centralizada**: las posiciones y reglas se cambian en un único objeto `CONFIG`, sin tocar la lógica.
- **Trazabilidad**: marca de tiempo automática en cada cambio de responsable.
- **Separación de responsabilidades**: el script mueve datos; las reglas de asignación viven en la hoja de control.
- **Reducción de errores manuales** en procesos repetitivos.

## Arquitectura

```text
Usuario edita una celda
        │
        ▼
 Trigger onEdit(e)
        │
        ├──► registerStatusTimestamp      Registra la fecha/hora del cambio (auditoría de eventos)
        │
        └──► syncActiveValueToControlSheet  Copia el tipo de solicitud a la hoja de control (LISTAS!N21)
                                │
                                ▼
                  Motor de reglas en la hoja LISTAS
                  (fórmulas: sugiere al funcionario con menor carga → LISTAS!O21)
                                │
                                ▼
        assignResponsibleFromRulesEngine    Copia el valor sugerido a la solicitud activa
```

## Estructura

```text
apps-script/Codigo.gs                       script (CONFIG + 3 funciones)
datos-ejemplo/SOLICITUDES_FICTICIAS.xlsx    libro de ejemplo (900 solicitudes ficticias)
herramientas/generar_libro_ficticio.py      generador del libro (semilla fija)
```

## Configuración central

```javascript
const CONFIG = {
  REQUESTS_SHEET: "SOLICITUDES",   // hoja de solicitudes
  CONTROL_SHEET: "LISTAS",         // hoja de control y reglas
  STATUS_COLUMN: 9,                // columna I: funcionario asignado (dispara la marca de tiempo)
  TIMESTAMP_OFFSET: 1,             // la marca de tiempo se escribe una columna a la derecha (J)
  MIN_DATA_ROW: 2,                 // primera fila de datos
  CONTEXT_MIN_ROW: 820,            // la sincronización aplica a filas posteriores
  CONTEXT_COLUMN: 3,               // columna C: tipo de solicitud
  CONTEXT_TARGET_CELL: "N21",      // destino en LISTAS
  ASSIGNMENT_SOURCE_CELL: "O21"    // resultado del motor de reglas
};
```

## Funciones

| Función | Disparador | Qué hace |
|---|---|---|
| `onEdit(e)` | Automático al editar | Orquesta las dos acciones siguientes |
| `registerStatusTimestamp(e)` | `onEdit` | Si se edita la columna de funcionario en `SOLICITUDES`, escribe la fecha y hora en la celda contigua |
| `syncActiveValueToControlSheet(e)` | `onEdit` | Si se edita el tipo de solicitud en una fila nueva (> 820), lo copia a `LISTAS!N21` |
| `assignResponsibleFromRulesEngine()` | Manual (menú o botón) | Copia `LISTAS!O21` a la celda activa de `SOLICITUDES` |

## Cómo probarlo
1. Importar [`SOLICITUDES_FICTICIAS.xlsx`](datos-ejemplo/SOLICITUDES_FICTICIAS.xlsx) en Google Sheets (*Archivo → Importar*).
2. *Extensiones → Apps Script*, pegar el contenido de [`apps-script/Codigo.gs`](apps-script/Codigo.gs) y guardar.
3. En `SOLICITUDES`, cambiar un funcionario en la columna **I**: la columna **J** recibe la fecha y hora.
4. En una fila posterior a la 820, editar la columna **C**: `LISTAS!N21` toma ese tipo y `LISTAS!O21` sugiere al funcionario con menor carga.
5. Seleccionar la celda de funcionario de una solicitud pendiente y ejecutar `assignResponsibleFromRulesEngine`.

## Datos de ejemplo
- **900 solicitudes** de 2024, 7 tipos de solicitud, 6 tipos de operador y 4 funcionarios ficticios.
- **20 solicitudes pendientes** de asignación (las últimas filas).
- Hoja `LISTAS` con la **carga por tipo y funcionario** (fórmulas `COUNTIFS`) y la regla de ejemplo de `O21`.

Carga resultante (referencia para verificar la regla):

| Tipo de solicitud | Func. A | Func. B | Func. C | Func. D | Sugerido |
|---|---:|---:|---:|---:|---|
| 01-AUTORIZACION | 37 | 49 | 53 | 52 | A |
| 02-RENOVACION | 61 | 68 | 63 | 67 | A |
| 03-CAMBIO DE DOMICILIO | 32 | 31 | 33 | 36 | B |
| 04-MODIFICACION | 31 | 28 | 23 | 29 | C |
| 05-CANCELACION | 6 | 13 | 17 | 11 | A |
| 06-CONSULTA | 15 | 19 | 15 | 24 | A |
| 07-OTROS | 18 | 19 | 13 | 17 | C |

> La regla de `O21` (menor carga) es un **ejemplo** para la demostración: el valor del proyecto está en
> que las reglas viven en la hoja y se pueden cambiar sin modificar el script.

## Alcance y limitaciones
- Usa un *trigger simple* `onEdit`: opera sobre el libro donde está instalado y no requiere autorización adicional.
- Está pensado para equipos pequeños: sin control de concurrencia ni bitácora histórica separada.
- Las posiciones de `CONFIG` dependen del diseño de la hoja; si cambia, se ajustan en un solo lugar.

## Stack
Google Apps Script (JavaScript) · Google Sheets · Python (openpyxl) para generar los datos de ejemplo

## Aviso
Proyecto de portafolio con fines demostrativos. No contiene información de organizaciones, personas ni
sistemas reales.
