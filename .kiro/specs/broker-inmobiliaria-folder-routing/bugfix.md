# Bugfix Requirements Document

## Introduction

Las funciones de búsqueda de archivos y gestión de carpetas en el flujo de renovación (`getPolicyFiles`, `obtenerCarpetaRenovacionPorPoliza`, `buscarCarpetaYGuardarArchivos`, `findDriveFolderByRef`) están hardcodeadas para buscar únicamente en la carpeta raíz de propietarios (ID: `1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC`). Esto causa que para los segmentos "Broker" e "Inmobiliaria" no se encuentren documentos ni se creen subcarpetas de renovación en la ubicación correcta, ya que sus archivos residen en carpetas individuales referenciadas en la columna AY de `SheetConsolidadoBrokersYInmobiliarias`.

## Bug Analysis

### Current Behavior (Defect)

1.1 WHEN the segment is "BROKER" or "INMOBILIARIA" AND `getPolicyFiles(policyRef)` is called THEN the system searches in the propietario root folder (ID `1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC`) and returns no files because broker/inmobiliaria documents are stored elsewhere

1.2 WHEN the segment is "BROKER" or "INMOBILIARIA" AND `obtenerCarpetaRenovacionPorPoliza(poliza, documento, asegurado)` is called THEN the system searches/creates folders under the propietario root folder instead of the broker/inmobiliaria case folder referenced in column AY of SheetConsolidadoBrokersYInmobiliarias

1.3 WHEN the segment is "BROKER" or "INMOBILIARIA" AND `buscarCarpetaYGuardarArchivos(archivosBase64, datos, ref)` is called THEN the system saves files to the propietario root folder structure instead of the correct broker/inmobiliaria folder

1.4 WHEN the segment is "BROKER" or "INMOBILIARIA" AND `findDriveFolderByRef(ref)` is called THEN the system searches only in the propietario root folder and fails to find the client folder because it does not exist there

### Expected Behavior (Correct)

2.1 WHEN the segment is "BROKER" or "INMOBILIARIA" AND `getPolicyFiles(policyRef)` is called THEN the system SHALL look up the policy number in columns BB (index 53) and BI (index 60) of `SheetConsolidadoBrokersYInmobiliarias`, retrieve the folder URL from column AY (index 50), and list files from that folder and its renovation subfolders

2.2 WHEN the segment is "BROKER" or "INMOBILIARIA" AND `obtenerCarpetaRenovacionPorPoliza(poliza, documento, asegurado)` is called THEN the system SHALL look up the policy in `SheetConsolidadoBrokersYInmobiliarias` (columns BB and BI), get the folder ID from column AY, open that folder, and search/create the "Renovacion" subfolder within it

2.3 WHEN the segment is "BROKER" or "INMOBILIARIA" AND `buscarCarpetaYGuardarArchivos(archivosBase64, datos, ref)` is called THEN the system SHALL resolve the correct broker/inmobiliaria folder from `SheetConsolidadoBrokersYInmobiliarias` column AY and save files to that folder's renovation subfolder

2.4 WHEN the segment is "BROKER" or "INMOBILIARIA" AND `findDriveFolderByRef(ref)` is called THEN the system SHALL resolve the folder from `SheetConsolidadoBrokersYInmobiliarias` column AY instead of searching the propietario root folder

### Unchanged Behavior (Regression Prevention)

3.1 WHEN the segment is "PROPIETARIO" AND `getPolicyFiles(policyRef)` is called THEN the system SHALL CONTINUE TO search in the propietario root folder (ID `1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC`) using `resolveClientReference` and `findDriveFolderByRef` as it does today

3.2 WHEN the segment is "PROPIETARIO" AND `obtenerCarpetaRenovacionPorPoliza(poliza, documento, asegurado)` is called THEN the system SHALL CONTINUE TO search/create folders under the propietario root folder using the existing reference lookup logic (SheetConsolidado → PolizasAntiguas → create new)

3.3 WHEN the segment is "PROPIETARIO" AND `buscarCarpetaYGuardarArchivos(archivosBase64, datos, ref)` is called THEN the system SHALL CONTINUE TO search and save files in the propietario root folder structure

3.4 WHEN the segment is "PROPIETARIO" AND `findDriveFolderByRef(ref)` is called THEN the system SHALL CONTINUE TO search in the propietario root folder by reference name

3.5 WHEN the segment is "BROKER" or "INMOBILIARIA" AND the policy is found in `SheetConsolidadoBrokersYInmobiliarias` THEN the system SHALL CONTINUE TO use the existing folder URL format from column AY (extracting the folder ID via `.split("/folders/")[1]`) as already done in `CargarPolizaBrokerYInmobiliaria` and `GenerarContratoBrokersInmobiliarias`

3.6 WHEN `normalizeSegmentoRenovacion_()` returns "PROPIETARIO" for empty or unrecognized segment values THEN the system SHALL CONTINUE TO default to the propietario flow, preserving backward compatibility for cases without explicit segment data

---

## Bug Condition (Formal)

```pascal
FUNCTION isBugCondition(X)
  INPUT: X of type RenovationFileRequest (contains segment, policyRef)
  OUTPUT: boolean
  
  // Returns true when the segment is BROKER or INMOBILIARIA
  RETURN normalizeSegmentoRenovacion_(X.dataLead, X.segmentoSheetColumn) IN {"BROKER", "INMOBILIARIA"}
END FUNCTION
```

```pascal
// Property: Fix Checking - Correct folder routing for Broker/Inmobiliaria
FOR ALL X WHERE isBugCondition(X) DO
  folder ← resolveFolder'(X)
  ASSERT folder.id = extractFolderId(SheetConsolidadoBrokersYInmobiliarias.getRange("AY" + findRow(X.poliza)).getDisplayValue())
  ASSERT folder IS NOT NULL
END FOR
```

```pascal
// Property: Preservation Checking - Propietario flow unchanged
FOR ALL X WHERE NOT isBugCondition(X) DO
  ASSERT resolveFolder(X) = resolveFolder'(X)
END FOR
```
