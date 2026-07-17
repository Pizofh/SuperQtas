# Operacion QTAS

## Modo seguro para produccion

El codigo ahora bloquea por defecto las operaciones mas destructivas.

Mientras `QTAS_ALLOW_DESTRUCTIVE` no este en `true`, quedan bloqueadas rutas como:

- `testResetEntornoQTAS()`
- reparaciones de headers que limpian hojas
- normalizacion masiva de `Precios_Referencia`

Recomendacion:

- En `prod`: deja `QTAS_ALLOW_DESTRUCTIVE` sin configurar.
- En `qa` o mantenimiento manual: activalo solo temporalmente si de verdad necesitas una operacion masiva.

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

## Carga historica sin alterar inventario

Si necesitas registrar compras, ventas o producciones atrasadas despues de haber contado el stock actual, pausa primero la sincronizacion de inventario. Los registros comerciales, costos, pagos y distribucion se guardan normalmente; solo se omiten sus movimientos de stock.

Desde el menu `QTAS ERP` del libro selecciona:

- `Pausar inventario para carga historica`
- Registra el historico pendiente.
- `Reanudar inventario operativo` cuando alcances el presente.

Tambien puedes ejecutar estas funciones desde Apps Script:

```javascript
pausarInventarioParaCargaHistoricaQTAS()
estadoSincronizacionInventarioQTAS()
reanudarInventarioOperativoQTAS()
```

La pausa no reconstruye ni modifica el snapshot existente. Las operaciones ingresadas durante ella quedan deliberadamente fuera del inventario para conservar el conteo fisico actual; las operaciones nuevas despues de reanudar vuelven a actualizarlo.

## Nota importante sobre QA

El archivo `QTAS_Testing.gs` sigue en el repo porque se usa para automatizacion, pero en `prod` ya no deberia poder vaciar hojas mientras `QTAS_ALLOW_DESTRUCTIVE` siga apagado.
