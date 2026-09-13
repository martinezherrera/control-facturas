# Control de facturas

Dashboard de control de ingreso de facturas.
Publicado en **https://martinezherrera.github.io/control-facturas/** — el enlace
no cambia; cada `push` actualiza lo que ven todos.

## Estructura

Una sola carpeta, que **es** el repositorio. Idéntica en cualquier PC.

| Archivo | Versionado | Rol |
|---|---|---|
| `index.html` | sí | La página. Estructura y estilos; no lleva datos. |
| `datos.js` | sí | Datos anonimizados. Se regenera en cada publicación. |
| `generar_dashboard_html.py` | sí | El generador. Lógica, sin datos. |
| `publicar.bat` | sí | Regenera y publica en un paso. |
| `control de ingresos facturas.xlsx` | **no** | La fuente. |
| `proveedores_map.csv` | **no** | Decodificador `PROV-nn` → razón social. |
| `dashboard_facturas.html` | **no** | Versión local, con razones sociales reales. |

Los tres últimos están bloqueados por `.gitignore` y **nunca** deben subirse:
el repositorio es público.

## Publicar una actualización

Actualiza la hoja `Exported` del Excel, guarda, y doble clic en `publicar.bat`.

Equivale a:

```
python generar_dashboard_html.py "control de ingresos facturas.xlsx" "dashboard_facturas.html" --web .
git add index.html datos.js
git commit -m "Actualizacion de datos"
git push
```

## Trabajar desde otro PC

```
git clone https://github.com/martinezherrera/control-facturas.git
```

El clon trae la página y la herramienta, pero **no** los datos. Copia a mano
dentro de la carpeta clonada:

- `control de ingresos facturas.xlsx`
- `proveedores_map.csv`

Requisitos: Python con `pandas` y `openpyxl`, y git.

### Por qué el mapa de proveedores es crítico

Los códigos `PROV-nn` se asignan por orden alfabético sobre los proveedores
presentes. Si corres el script sin `proveedores_map.csv`, se genera uno desde
cero: mientras el conjunto de proveedores sea idéntico da el mismo resultado,
pero en cuanto aparezca un proveedor nuevo que caiga al medio del alfabeto,
todos los códigos posteriores se corren. Quien tenga la tabla anterior leería
mal el dashboard sin notarlo.

Ese archivo viaja siempre con el Excel.

## Qué queda visible públicamente

Montos facturados, cierres de mes, presupuestos mensuales y número de órdenes
por canal. Los proveedores aparecen solo como código.

Es una decisión tomada a conciencia. Si deja de ser aceptable, la salida es
repositorio privado con un plan que habilite Pages (Pro, Team o Enterprise);
en plan gratuito, Pages solo funciona desde repositorios públicos.
