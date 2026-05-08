# Changelog

Todos los cambios notables de este proyecto se documentan en este archivo.

El formato está basado en [Keep a Changelog](https://keepachangelog.com/es-ES/1.1.0/)
y este proyecto adhiere a [Versionamiento Semántico](https://semver.org/lang/es/).

## [No publicado]

### Corregido
- Se corrigió el routing de carpetas en el flujo de renovación para segmentos Broker e Inmobiliaria: las funciones `getPolicyFiles`, `obtenerCarpetaRenovacionPorPoliza`, `buscarCarpetaYGuardarArchivos` ahora consultan el segmento y buscan la carpeta correcta desde la columna AY de `SheetConsolidadoBrokersYInmobiliarias` en lugar de buscar en la carpeta raíz de propietarios

### Cambiado
- Se modificó la llamada a `getPolicyFiles` desde el frontend para pasar el segmento como segundo parámetro, permitiendo el routing correcto sin búsquedas innecesarias

### Agregado
- Se creó función helper `resolverCarpetaBrokerInmobiliaria_(poliza)` que busca la póliza en columnas BI y BB de la hoja de Brokers/Inmobiliarias y retorna la carpeta de Drive correspondiente desde columna AY
