# Broker/Inmobiliaria Folder Routing — Bugfix Design

## Overview

Las funciones de gestión de carpetas en el flujo de renovación (`getPolicyFiles`, `obtenerCarpetaRenovacionPorPoliza`, `buscarCarpetaYGuardarArchivos`, `findDriveFolderByRef`) están hardcodeadas para buscar únicamente en la carpeta raíz de propietarios (ID: `1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC`). Para los segmentos "Broker" e "Inmobiliaria", la carpeta correcta se encuentra referenciada en la columna AY de `SheetConsolidadoBrokersYInmobiliarias`, y la póliza se identifica en las columnas BB (index 53, observaciones) y BI (index 60, número de póliza emitida).

La estrategia de corrección consiste en crear una función helper compartida (`resolverCarpetaBrokerInmobiliaria_`) que busque la póliza en la hoja de Brokers/Inmobiliarias y devuelva la carpeta correspondiente. Cada función afectada incorporará un routing condicional: si el segmento es "PROPIETARIO" → lógica actual; si es "BROKER" o "INMOBILIARIA" → usar el helper.

## Glossary

- **Bug_Condition (C)**: El segmento normalizado es "BROKER" o "INMOBILIARIA" — las funciones de carpeta buscan en la ubicación incorrecta (carpeta raíz de propietarios)
- **Property (P)**: Para segmentos Broker/Inmobiliaria, las funciones de carpeta deben resolver la carpeta desde la columna AY de `SheetConsolidadoBrokersYInmobiliarias`
- **Preservation**: El flujo de propietarios (búsqueda en carpeta raíz `1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC`) debe permanecer intacto
- **`normalizeSegmentoRenovacion_`**: Función en `Código.js` que normaliza el segmento a "PROPIETARIO", "BROKER" o "INMOBILIARIA"
- **`SheetConsolidadoBrokersYInmobiliarias`**: Hoja "Consolidado" del spreadsheet `1IBl08te0QJ27DYU9qGZ3qVRlOowQxWbc1Cu_Eu32g_8` que contiene datos de casos Broker/Inmobiliaria
- **Columna AY (index 50)**: URL de la carpeta de Drive del caso Broker/Inmobiliaria (formato: `https://drive.google.com/.../folders/{folderId}`)
- **Columna BB (index 53)**: Observaciones/texto de referencia de póliza (usado por `EnviarCorreccionDocsBrokersInmobiliarias`)
- **Columna BI (index 60)**: Número de póliza emitida (escrito por `CargarPolizaBrokerYInmobiliaria` tras OCR)
- **Columna C (index 2)**: Código de solicitud/referencia del caso (usado como lookup principal)

## Bug Details

### Bug Condition

El bug se manifiesta cuando el segmento del lead es "BROKER" o "INMOBILIARIA" y se invoca cualquiera de las funciones de gestión de carpetas del flujo de renovación. Estas funciones usan exclusivamente el `rootId = "1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC"` (carpeta de propietarios) para buscar/crear carpetas, ignorando que los casos Broker/Inmobiliaria tienen su propia carpeta referenciada en la columna AY.

**Formal Specification:**
```
FUNCTION isBugCondition(input)
  INPUT: input of type {poliza: string, segmento: string, dataLead: object}
  OUTPUT: boolean
  
  segmentoNorm := normalizeSegmentoRenovacion_(input.dataLead, input.segmentoSheetColumn)
  RETURN segmentoNorm IN {"BROKER", "INMOBILIARIA"}
         AND functionCalled IN {"getPolicyFiles", "obtenerCarpetaRenovacionPorPoliza", 
                                "buscarCarpetaYGuardarArchivos", "findDriveFolderByRef"}
         AND folderResolution uses rootId "1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC"
END FUNCTION
```

### Examples

- **Ejemplo 1**: Un lead Broker con póliza "POL-12345" tiene su carpeta en `https://drive.google.com/drive/folders/abc123` (columna AY). Al llamar `getPolicyFiles("POL-12345")`, el sistema busca en la carpeta raíz de propietarios y retorna `{success: false, files: []}`. **Esperado**: buscar en la carpeta `abc123` y retornar los archivos encontrados.
- **Ejemplo 2**: `processAnalystDecision` con `payload.segmento = "BROKER"` llama a `obtenerCarpetaRenovacionPorPoliza("POL-12345", "CC-999", "Juan")`. El sistema busca/crea carpeta bajo la raíz de propietarios. **Esperado**: obtener la carpeta desde columna AY y buscar/crear subcarpeta "Renovacion" dentro de ella.
- **Ejemplo 3**: `buscarCarpetaYGuardarArchivos(archivos, datos, "REF-001")` para un caso Inmobiliaria guarda archivos en la carpeta de propietarios. **Esperado**: resolver la carpeta desde columna AY y guardar allí.
- **Ejemplo 4 (edge case)**: Un lead con segmento vacío o "Sin segmento" → `normalizeSegmentoRenovacion_` retorna "PROPIETARIO" → debe seguir usando la carpeta raíz actual (no es bug condition).

## Expected Behavior

### Preservation Requirements

**Unchanged Behaviors:**
- El flujo completo de propietarios (búsqueda por referencia en `SheetConsolidado` → `PolizasAntiguas` → creación de carpeta nueva) debe continuar funcionando exactamente igual
- La función `resolveClientReference` no se modifica — sigue buscando en `SheetConsolidado` y `PolizasAntiguas`
- El patrón de extracción de folder ID desde URL (`.split("/folders/")[1]`) se reutiliza tal como existe en `CargarPolizaBrokerYInmobiliaria` y `GenerarContratoBrokersInmobiliarias`
- La función `normalizeSegmentoRenovacion_` no se modifica
- La lógica de `guardarArchivosEnCarpeta` (creación de blobs, naming, subcarpeta "Procesos Especiales") permanece intacta
- El frontend (`Main.JS.html`) no requiere cambios — las interfaces de las funciones expuestas al frontend se mantienen compatibles

**Scope:**
Todas las invocaciones donde `normalizeSegmentoRenovacion_` retorna "PROPIETARIO" deben ser completamente inalteradas por este fix. Esto incluye:
- Leads sin segmento explícito (default a PROPIETARIO)
- Leads con segmento "Propietario", "propietario", "Sin segmento", etc.
- Cualquier llamada a `getPolicyFiles` donde la póliza NO existe en `SheetConsolidadoBrokersYInmobiliarias`

## Hypothesized Root Cause

Basado en el análisis del código, la causa raíz es clara y directa:

1. **Hardcoding del rootId**: Las cuatro funciones afectadas tienen `const rootId = "1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC"` hardcodeado como punto de partida para toda búsqueda de carpetas. Este ID corresponde exclusivamente a la carpeta raíz de propietarios.

2. **Ausencia de routing por segmento**: Cuando se implementaron estas funciones, solo existía el flujo de propietarios. El flujo de Broker/Inmobiliaria se agregó después (evidenciado por funciones como `CargarPolizaBrokerYInmobiliaria` que SÍ usan la columna AY), pero las funciones de renovación nunca se actualizaron.

3. **`getPolicyFiles` no recibe segmento**: Esta función es llamada desde el frontend con solo `policyRef` como parámetro. No tiene forma de saber el segmento sin hacer una búsqueda interna en `SheetConsolidadoBrokersYInmobiliarias`.

4. **`obtenerCarpetaRenovacionPorPoliza` no recibe segmento**: Aunque es llamada desde `processAnalystDecision` (que SÍ tiene `payload.segmento` y `payload.dataLead`), la función no acepta un parámetro de segmento.

## Correctness Properties

Property 1: Bug Condition - Folder Resolution para Broker/Inmobiliaria

_For any_ input donde el segmento normalizado es "BROKER" o "INMOBILIARIA" y la póliza existe en `SheetConsolidadoBrokersYInmobiliarias` (columnas BB o BI), las funciones de carpeta corregidas SHALL resolver la carpeta de Drive desde la columna AY de esa misma fila, extrayendo el folder ID con `.split("/folders/")[1]`, y operar sobre esa carpeta (listar archivos, buscar/crear subcarpeta "Renovacion", guardar archivos).

**Validates: Requirements 2.1, 2.2, 2.3, 2.4**

Property 2: Preservation - Flujo Propietario Inalterado

_For any_ input donde el segmento normalizado es "PROPIETARIO" (incluyendo segmentos vacíos o no reconocidos), las funciones de carpeta corregidas SHALL producir exactamente el mismo resultado que las funciones originales, preservando la búsqueda en la carpeta raíz `1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC` mediante `resolveClientReference` y búsqueda por nombre de referencia.

**Validates: Requirements 3.1, 3.2, 3.3, 3.4, 3.5, 3.6**

## Fix Implementation

### Changes Required

Asumiendo que nuestro análisis de causa raíz es correcto:

**File**: `Código.js`

**Specific Changes**:

1. **Crear helper `resolverCarpetaBrokerInmobiliaria_`** (función privada compartida):
   - Recibe `poliza` (string) como parámetro
   - Busca la póliza en columna BI (index 60) de `SheetConsolidadoBrokersYInmobiliarias` usando `createTextFinder`
   - Si no encuentra, busca en columna BB (index 53) con `createTextFinder` (búsqueda parcial, ya que BB contiene texto de observaciones)
   - Si encuentra una fila, extrae la URL de columna AY de esa fila
   - Extrae el folder ID con `.split("/folders/")[1]`
   - Retorna `{found: true, folder: DriveApp.getFolderById(folderId), fila: rowNumber}` o `{found: false}`
   - Patrón idéntico al ya usado en `CargarPolizaBrokerYInmobiliaria`

2. **Modificar `getPolicyFiles(policyRef)`**:
   - Antes de la lógica actual, intentar resolver con `resolverCarpetaBrokerInmobiliaria_(policyRef)`
   - Si `found === true` → usar esa carpeta como `clientFolder` y continuar con el escaneo de archivos existente
   - Si `found === false` → continuar con la lógica actual (propietario: `resolveClientReference` + `findDriveFolderByRef`)
   - No se modifica la firma de la función (sigue recibiendo solo `policyRef`)

3. **Modificar `obtenerCarpetaRenovacionPorPoliza(poliza, documento, asegurado)`** — agregar parámetro opcional `segmento`:
   - Nueva firma: `obtenerCarpetaRenovacionPorPoliza(poliza, documento, asegurado, segmento)`
   - Si `segmento` es "BROKER" o "INMOBILIARIA" (normalizado), usar `resolverCarpetaBrokerInmobiliaria_(poliza)`
   - Si el helper retorna `found === true`, buscar/crear subcarpeta "Renovacion" dentro de esa carpeta
   - Si `segmento` es "PROPIETARIO" o no se proporciona → lógica actual intacta

4. **Modificar la llamada en `processAnalystDecision`**:
   - Pasar el segmento disponible: `obtenerCarpetaRenovacionPorPoliza(poliza, documento, asegurado, payload.segmento || (payload.dataLead && payload.dataLead.segmento))`

5. **Modificar `buscarCarpetaYGuardarArchivos(archivosBase64, datos, ref)`** — agregar parámetro opcional `segmento`:
   - Nueva firma: `buscarCarpetaYGuardarArchivos(archivosBase64, datos, ref, segmento)`
   - Si segmento normalizado es "BROKER" o "INMOBILIARIA", usar `resolverCarpetaBrokerInmobiliaria_` con `datos.poliza` para obtener la carpeta
   - Si es "PROPIETARIO" o no se proporciona → lógica actual intacta

6. **Modificar `findDriveFolderByRef(ref)`** — agregar parámetro opcional `segmento` y `poliza`:
   - Nueva firma: `findDriveFolderByRef(ref, segmento, poliza)`
   - Si segmento normalizado es "BROKER" o "INMOBILIARIA" y `poliza` está disponible, usar `resolverCarpetaBrokerInmobiliaria_(poliza)`
   - Si es "PROPIETARIO" o parámetros no proporcionados → lógica actual intacta
   - La llamada desde `getPolicyFiles` ya se maneja internamente (punto 2), por lo que `findDriveFolderByRef` puede mantenerse sin cambios si `getPolicyFiles` hace el routing antes de llamarla

## Testing Strategy

### Validation Approach

La estrategia de testing sigue un enfoque de dos fases: primero, demostrar el bug con contraejemplos en el código sin corregir; luego, verificar que el fix funciona correctamente y preserva el comportamiento existente.

### Exploratory Bug Condition Checking

**Goal**: Demostrar contraejemplos que evidencien el bug ANTES de implementar el fix. Confirmar o refutar el análisis de causa raíz.

**Test Plan**: Simular llamadas a las funciones afectadas con pólizas que existen en `SheetConsolidadoBrokersYInmobiliarias` y verificar que el sistema busca en la carpeta incorrecta. Ejecutar en código SIN corregir para observar fallos.

**Test Cases**:
1. **getPolicyFiles con póliza Broker**: Llamar `getPolicyFiles("POL-BROKER-001")` donde la póliza existe en columna BI de la hoja de Brokers → retorna `{success: false}` porque busca en carpeta de propietarios (fallará en código sin corregir)
2. **obtenerCarpetaRenovacionPorPoliza con caso Broker**: Llamar con una póliza que solo existe en la hoja de Brokers → crea carpeta nueva bajo propietarios en vez de usar la existente (fallará en código sin corregir)
3. **findDriveFolderByRef con referencia Broker**: Buscar una referencia que no existe como subcarpeta en la raíz de propietarios pero sí tiene carpeta en columna AY → retorna `null` (fallará en código sin corregir)

**Expected Counterexamples**:
- `getPolicyFiles` retorna `{success: false, files: []}` para pólizas Broker/Inmobiliaria válidas
- `obtenerCarpetaRenovacionPorPoliza` crea carpetas duplicadas bajo la raíz de propietarios
- Causa confirmada: hardcoding de `rootId` sin routing por segmento

### Fix Checking

**Goal**: Verificar que para todas las entradas donde la bug condition se cumple, las funciones corregidas producen el comportamiento esperado.

**Pseudocode:**
```
FOR ALL input WHERE isBugCondition(input) DO
  result := functionFixed(input)
  ASSERT result.folder.getId() = extractFolderId(
    SheetConsolidadoBrokersYInmobiliarias.getRange("AY" + findRow(input.poliza)).getDisplayValue()
  )
  ASSERT result.folder IS NOT NULL
  ASSERT result.success = true (para getPolicyFiles)
END FOR
```

### Preservation Checking

**Goal**: Verificar que para todas las entradas donde la bug condition NO se cumple, las funciones corregidas producen el mismo resultado que las funciones originales.

**Pseudocode:**
```
FOR ALL input WHERE NOT isBugCondition(input) DO
  ASSERT getPolicyFiles_original(input.poliza) = getPolicyFiles_fixed(input.poliza)
  ASSERT obtenerCarpeta_original(input.poliza, input.doc, input.aseg) 
       = obtenerCarpeta_fixed(input.poliza, input.doc, input.aseg, "PROPIETARIO")
  ASSERT findDriveFolder_original(input.ref) = findDriveFolder_fixed(input.ref)
END FOR
```

**Testing Approach**: Property-based testing es recomendado para preservation checking porque:
- Genera muchos casos de prueba automáticamente sobre el dominio de entradas propietario
- Detecta edge cases que tests manuales podrían omitir (segmentos vacíos, null, "Sin segmento", etc.)
- Provee garantías fuertes de que el comportamiento no cambia para entradas no-buggy

**Test Plan**: Observar el comportamiento en código SIN corregir para entradas de propietario, luego escribir tests que capturen ese comportamiento.

**Test Cases**:
1. **Preservación getPolicyFiles propietario**: Verificar que pólizas de propietarios siguen encontrando archivos en la carpeta raíz
2. **Preservación obtenerCarpetaRenovacionPorPoliza propietario**: Verificar que la lógica de búsqueda SheetConsolidado → PolizasAntiguas → crear nueva sigue funcionando
3. **Preservación segmento vacío/null**: Verificar que `normalizeSegmentoRenovacion_` con valores vacíos retorna "PROPIETARIO" y el flujo no se altera
4. **Preservación findDriveFolderByRef sin parámetros extra**: Verificar que llamadas sin segmento siguen buscando en la raíz de propietarios

### Unit Tests

- Test `resolverCarpetaBrokerInmobiliaria_` con póliza existente en columna BI → retorna folder correcto
- Test `resolverCarpetaBrokerInmobiliaria_` con póliza existente en columna BB → retorna folder correcto
- Test `resolverCarpetaBrokerInmobiliaria_` con póliza inexistente → retorna `{found: false}`
- Test `getPolicyFiles` con póliza Broker → lista archivos de la carpeta correcta
- Test `obtenerCarpetaRenovacionPorPoliza` con segmento "BROKER" → usa carpeta de columna AY
- Test `obtenerCarpetaRenovacionPorPoliza` con segmento null → usa lógica propietario
- Test edge case: columna AY vacía o con URL malformada → manejo graceful de error

### Property-Based Tests

- Generar segmentos aleatorios y verificar que el routing es correcto (BROKER/INMOBILIARIA → helper, PROPIETARIO/vacío/null → raíz)
- Generar variaciones de `normalizeSegmentoRenovacion_` inputs y verificar que el default siempre es PROPIETARIO
- Verificar que para cualquier póliza de propietario, el resultado de las funciones corregidas es idéntico al de las originales

### Integration Tests

- Test flujo completo: `processAnalystDecision` con payload Broker → archivos guardados en carpeta correcta
- Test flujo completo: `processAnalystDecision` con payload Propietario → archivos guardados en carpeta raíz (sin cambios)
- Test `getPolicyFiles` llamado desde frontend con póliza Broker → retorna archivos de la carpeta AY
- Test que la subcarpeta "Renovacion" se crea correctamente dentro de la carpeta Broker/Inmobiliaria
