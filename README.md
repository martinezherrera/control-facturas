# Control de facturas

Dashboard de control de ingreso de facturas. Se publica con GitHub Pages y el
enlace no cambia: cada `push` actualiza lo que ven todos.

## Qué contiene

| Archivo | Rol |
|---|---|
| `index.html` | La página. Estructura y estilos; no lleva datos. Cambia poco. |
| `datos.js` | Los datos, regenerados desde el Excel en cada publicación. |

## Qué NO contiene, y por qué

Este repositorio es **público**. Por eso:

- El Excel de origen no se versiona.
- Los proveedores aparecen como `PROV-01`, `PROV-02`, … La correspondencia con
  las razones sociales vive en `proveedores_map.csv`, junto al Excel, **fuera de
  este repositorio**. El `.gitignore` la bloquea.
- La versión con nombres reales es `dashboard_facturas.html`, que se queda en el
  computador y tampoco se versiona.

Lo que sí queda visible para cualquiera con la URL: montos facturados, cierres
de mes, presupuestos mensuales y número de órdenes por canal. Es una decisión
tomada a conciencia, no un descuido; si en algún momento deja de ser aceptable,
la salida es pasar el repositorio a privado con un plan que habilite Pages.

## Cómo publicar una actualización

Desde `C:\CONTROL FACTURAS`, después de actualizar la hoja `Exported`:

```
publicar.bat
```

Eso regenera `datos.js` y hace `commit` y `push`. El sitio tarda entre 30
segundos y un par de minutos en reflejar el cambio.

## Códigos de proveedor

Son estables: un proveedor conserva su código entre ejecuciones. Los nuevos
reciben el siguiente correlativo libre. No reordenar ni editar
`proveedores_map.csv` a mano, o los códigos publicados dejarán de coincidir con
los que ya circulan.
