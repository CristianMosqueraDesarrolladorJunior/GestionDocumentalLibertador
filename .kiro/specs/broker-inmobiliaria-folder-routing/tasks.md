# Implementation Plan

- [x] 1. Write bug condition exploration test
  - **Property 1: Bug Condition** - Broker/Inmobiliaria Folder Resolution Fails
  - **CRITICAL**: This test MUST FAIL on unfixed code - failure confirms the bug exists
  - **DO NOT attempt to fix the test or the code when it fails**
  - **NOTE**: This test encodes the expected behavior - it will validate the fix when it passes after implementation
  - **GOAL**: Surface counterexamples that demonstrate the bug exists
  - **Scoped PBT Approach**: Scope the property to concrete failing cases — call `getPolicyFiles` with a policy number that exists in `SheetConsolidadoBrokersYInmobiliarias` column BI (index 60) and verify it resolves the folder from column AY
  - **Manual Verification Steps** (Google Apps Script — no automated test framework):
    1. Identify a real policy number in `SheetConsolidadoBrokersYInmobiliarias` column BI (e.g., row with folder URL in column AY)
    2. Call `getPolicyFiles("<broker-policy-number>")` from the Apps Script editor
    3. Observe: returns `{success: false, files: []}` because it searches in propietario root folder `1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC`
    4. Call `obtenerCarpetaRenovacionPorPoliza("<broker-policy>", "doc", "asegurado")` — observe it creates/searches under propietario root instead of the broker folder from column AY
    5. Call `findDriveFolderByRef("<broker-ref>")` — observe it returns `null` because the folder doesn't exist under propietario root
  - **EXPECTED OUTCOME**: All calls fail to find the correct folder (confirms bug exists — hardcoded `rootId` ignores segment)
  - Document counterexamples: e.g., `getPolicyFiles("POL-BROKER-001")` returns `{success: false}` instead of listing files from the folder in column AY
  - Mark task complete when verification is executed and failure is documented
  - _Requirements: 1.1, 1.2, 1.3, 1.4_

- [x] 2. Write preservation property tests (BEFORE implementing fix)
  - **Property 2: Preservation** - Propietario Flow Unchanged
  - **IMPORTANT**: Follow observation-first methodology
  - **Manual Verification Steps** (Google Apps Script — no automated test framework):
    1. Identify a real propietario policy number in `SheetConsolidado` column BE
    2. Call `getPolicyFiles("<propietario-policy>")` — observe and record the result (files found in root folder `1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC`)
    3. Call `obtenerCarpetaRenovacionPorPoliza("<propietario-policy>", "doc", "asegurado")` — observe and record the folder URL returned
    4. Call `findDriveFolderByRef("<propietario-ref>")` — observe and record the folder object returned
    5. Test with empty/null segment: verify `normalizeSegmentoRenovacion_({}, null)` returns "PROPIETARIO"
    6. Test with "Sin segmento": verify `normalizeSegmentoRenovacion_({segmento: "Sin segmento"}, null)` returns "PROPIETARIO"
  - **Preservation Property**: For all inputs where `normalizeSegmentoRenovacion_` returns "PROPIETARIO", the functions must continue using root folder `1e05FPKAfrRnqBbUpOF1JP9ostCg2TsjC` and the existing `resolveClientReference` + `findDriveFolderByRef` logic
  - **EXPECTED OUTCOME**: All propietario calls work correctly on UNFIXED code (confirms baseline behavior to preserve)
  - Record observed outputs as baseline for post-fix comparison
  - Mark task complete when observations are recorded and all propietario calls pass
  - _Requirements: 3.1, 3.2, 3.3, 3.4, 3.5, 3.6_

- [x] 3. Fix for Broker/Inmobiliaria folder routing

  - [x] 3.1 Create helper function `resolverCarpetaBrokerInmobiliaria_(poliza)`
    - Add new private function in `Código.js` after `findDriveFolderByRef`
    - Receives `poliza` (string) as parameter
    - Search policy in column BI (index 60) of `SheetConsolidadoBrokersYInmobiliarias` using `createTextFinder` with `matchEntireCell(true)`
    - If not found in BI, search in column BB (index 53) using `createTextFinder` (partial match — BB contains observation text)
    - If a row is found, extract folder URL from column AY (index 50) of that row
    - Extract folder ID with `.split("/folders/")[1]`
    - Return `{found: true, folder: DriveApp.getFolderById(folderId), fila: rowNumber}` or `{found: false}`
    - Pattern identical to existing `CargarPolizaBrokerYInmobiliaria` folder resolution
    - Include JSDoc with parameter and return type documentation
    - _Bug_Condition: isBugCondition(input) where normalizeSegmentoRenovacion_ IN {"BROKER", "INMOBILIARIA"}_
    - _Expected_Behavior: resolverCarpetaBrokerInmobiliaria_ returns the folder from column AY of SheetConsolidadoBrokersYInmobiliarias_
    - _Preservation: Function only called for BROKER/INMOBILIARIA segments; PROPIETARIO flow never invokes it_
    - _Requirements: 2.1, 2.2, 2.3, 2.4_

  - [x] 3.2 Modify `getPolicyFiles(policyRef)` to add segment-aware routing
    - Before the current `resolveClientReference` logic, call `resolverCarpetaBrokerInmobiliaria_(policyRef)`
    - If `found === true` → use that folder as `clientFolder` and skip `resolveClientReference` + `findDriveFolderByRef`
    - If `found === false` → continue with current propietario logic unchanged
    - Do NOT modify the function signature (still receives only `policyRef`)
    - The function auto-detects broker/inmobiliaria by checking if the policy exists in `SheetConsolidadoBrokersYInmobiliarias`
    - _Bug_Condition: policyRef exists in SheetConsolidadoBrokersYInmobiliarias columns BI or BB_
    - _Expected_Behavior: getPolicyFiles resolves folder from column AY and lists files from it_
    - _Preservation: If policy not found in SheetConsolidadoBrokersYInmobiliarias, falls through to existing propietario logic_
    - _Requirements: 2.1, 3.1_

  - [x] 3.3 Modify `obtenerCarpetaRenovacionPorPoliza` to accept optional `segmento` parameter
    - Change signature to `obtenerCarpetaRenovacionPorPoliza(poliza, documento, asegurado, segmento)`
    - At the start of the function, normalize segmento: `var segNorm = (segmento || "").toUpperCase().trim()`
    - If `segNorm === "BROKER" || segNorm === "INMOBILIARIA"` → call `resolverCarpetaBrokerInmobiliaria_(poliza)`
    - If helper returns `found === true` → search/create "Renovacion" subfolder within that folder (reuse existing subfolder search pattern)
    - If helper returns `found === false` → fall through to current propietario logic
    - If `segNorm` is "PROPIETARIO" or empty/not provided → current logic unchanged (backward compatible)
    - _Bug_Condition: segmento normalized is "BROKER" or "INMOBILIARIA"_
    - _Expected_Behavior: Uses folder from column AY and creates/finds "Renovacion" subfolder within it_
    - _Preservation: When segmento is empty/null/"PROPIETARIO", existing rootId-based logic executes unchanged_
    - _Requirements: 2.2, 3.2_

  - [x] 3.4 Modify call in `processAnalystDecision` to pass segment
    - Change the call from `obtenerCarpetaRenovacionPorPoliza(poliza, documento, asegurado)` to `obtenerCarpetaRenovacionPorPoliza(poliza, documento, asegurado, payload.segmento || (payload.dataLead && payload.dataLead.segmento))`
    - This passes the segment information available in the payload to enable routing
    - _Bug_Condition: processAnalystDecision has segment info but doesn't pass it_
    - _Expected_Behavior: Segment is forwarded to obtenerCarpetaRenovacionPorPoliza for routing_
    - _Preservation: If payload.segmento and dataLead.segmento are both undefined/null, passes undefined → function defaults to propietario_
    - _Requirements: 2.2, 3.2_

  - [x] 3.5 Modify `buscarCarpetaYGuardarArchivos` to accept optional `segmento` parameter
    - Change signature to `buscarCarpetaYGuardarArchivos(archivosBase64, datos, ref, segmento)`
    - At the start, normalize segmento: `var segNorm = (segmento || "").toUpperCase().trim()`
    - If `segNorm === "BROKER" || segNorm === "INMOBILIARIA"` → call `resolverCarpetaBrokerInmobiliaria_(datos.poliza)`
    - If helper returns `found === true` → use that folder as `carpetaCliente` and search/create "Renovacion" subfolder within it
    - If helper returns `found === false` → fall through to current propietario logic (search in rootId)
    - If `segNorm` is "PROPIETARIO" or empty/not provided → current logic unchanged
    - _Bug_Condition: segmento normalized is "BROKER" or "INMOBILIARIA"_
    - _Expected_Behavior: Resolves folder from column AY and saves files there_
    - _Preservation: When segmento is empty/null/"PROPIETARIO", existing rootId search logic executes unchanged_
    - _Requirements: 2.3, 3.3_

  - [x] 3.6 Verify bug condition exploration test now passes
    - **Property 1: Expected Behavior** - Broker/Inmobiliaria Folder Resolution Works
    - **IMPORTANT**: Re-run the SAME verification from task 1 - do NOT write a new test
    - Re-execute the manual verification steps from task 1 on the FIXED code:
      1. Call `getPolicyFiles("<broker-policy-number>")` → should now return `{success: true, files: [...]}` with files from the column AY folder
      2. Call `obtenerCarpetaRenovacionPorPoliza("<broker-policy>", "doc", "asegurado", "BROKER")` → should return folder within the broker's AY folder
    - **EXPECTED OUTCOME**: All calls now resolve the correct broker/inmobiliaria folder (confirms bug is fixed)
    - _Requirements: 2.1, 2.2, 2.3_

  - [x] 3.7 Verify preservation tests still pass
    - **Property 2: Preservation** - Propietario Flow Still Works After Fix
    - **IMPORTANT**: Re-run the SAME verification from task 2 - do NOT write new tests
    - Re-execute the manual verification steps from task 2 on the FIXED code:
      1. Call `getPolicyFiles("<propietario-policy>")` → should return same result as baseline
      2. Call `obtenerCarpetaRenovacionPorPoliza("<propietario-policy>", "doc", "asegurado")` → should return same folder URL as baseline
      3. Call `findDriveFolderByRef("<propietario-ref>")` → should return same folder as baseline
      4. Verify `normalizeSegmentoRenovacion_` still returns "PROPIETARIO" for empty/null/"Sin segmento"
    - **EXPECTED OUTCOME**: All propietario calls produce identical results to baseline (confirms no regressions)
    - Confirm all observations match the baseline recorded in task 2

- [x] 4. Checkpoint - Ensure all tests pass
  - Verify `getPolicyFiles` works for both Broker/Inmobiliaria AND Propietario policies
  - Verify `obtenerCarpetaRenovacionPorPoliza` routes correctly based on segment parameter
  - Verify `buscarCarpetaYGuardarArchivos` routes correctly based on segment parameter
  - Verify `processAnalystDecision` passes segment to `obtenerCarpetaRenovacionPorPoliza`
  - Verify no changes to `normalizeSegmentoRenovacion_`, `resolveClientReference`, or `guardarArchivosEnCarpeta`
  - Verify frontend (`Main.JS.html`) requires no changes — function signatures exposed to frontend remain compatible
  - Ensure all tests pass, ask the user if questions arise.
