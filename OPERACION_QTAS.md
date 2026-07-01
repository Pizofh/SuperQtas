# Operacion QTAS

## Modo seguro para produccion

El codigo ahora bloquea por defecto las operaciones mas destructivas.

Mientras `QTAS_ALLOW_DESTRUCTIVE` no este en `true`, quedan bloqueadas rutas como:

- `testResetEntornoQTAS()`
- `testLimpiarRegistrosPruebaQTAS()`
- `recalcularSaldosQTAS()`
- `reconstruirDistribucionIngresosQTAS_()`
- reparaciones de headers que limpian hojas
- normalizacion masiva de `Precios_Referencia`

Recomendacion:

- En `prod`: deja `QTAS_ALLOW_DESTRUCTIVE` sin configurar.
- En `qa` o mantenimiento manual: activalo solo temporalmente si de verdad necesitas una operacion masiva.

Funciones utiles:

- `habilitarOperacionesDestructivasQTAS()`
- `bloquearOperacionesDestructivasQTAS()`

## Backups diarios

Funciones nuevas:

- `instalarBackupDiarioQTAS({ hour: 3 })`
- `ejecutarBackupDiarioQTAS()`
- `crearBackupManualQTAS()`
- `getEstadoBackupsQTAS()`
- `desinstalarBackupDiarioQTAS()`

### Instalacion recomendada

En el proyecto de Apps Script que vas a usar como `prod`, corre una vez:

```javascript
instalarBackupDiarioQTAS({ hour: 3 })
```

Eso crea un trigger diario y guarda copias del spreadsheet en una carpeta `QTAS_Backups` junto al archivo principal.

## Politica de retencion

El backup ahora mantiene solo un archivo activo por spreadsheet.

Cada vez que corre:

- reutiliza el mismo archivo de backup
- lo refresca con el estado actual del spreadsheet

Eso aplica tanto para:

- `ejecutarBackupDiarioQTAS()`
- `crearBackupManualQTAS()`

### Verificar estado

```javascript
getEstadoBackupsQTAS()
```

### Crear backup manual antes de tocar algo delicado

```javascript
crearBackupManualQTAS()
```

## Nota importante sobre QA

El archivo `QTAS_Testing.gs` sigue en el repo porque se usa para automatizacion, pero en `prod` ya no deberia poder vaciar hojas mientras `QTAS_ALLOW_DESTRUCTIVE` siga apagado.
