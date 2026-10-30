# Dueñas — Google Apps Script

## Reglas críticas
- **NUNCA tocar** `gestion-almacen/mojados/` — funciona perfecto, no modificar.
- Push siempre desde este directorio: `npx clasp push --force`
- Spreadsheet ID: `1QDXElIotoCZEErCAbabJgdYaYmO-eUXI8eD18Es-O2M`

## Estructura de columnas (estado actual)

### ENTRADAS (fila 1 = título, fila 2 = cabecera, datos desde fila 3)
| Col | A | B | C | D | E | F | G | H |
|-----|---|---|---|---|---|---|---|---|
| Campo | Nº Albarán | Carta Porte | Fecha | Referencia | Piezas/palet | Nº Palets | Total piezas | Cliente |
| Índice | 0 | 1 | 2 | 3 | 4 | 5 | 6 | 7 |

Detección formato: cabecera[1] contiene "carta" → nueva (8 cols). Si no → antigua (7 cols, sin CMR).

### SALIDAS (fila 1 = título, fila 2 = cabecera, datos desde fila 3)
| Col | A | B | C | D | E | F |
|-----|---|---|---|---|---|---|
| Campo | Fecha | Referencia | Piezas/palet | Nº Palets | Total piezas | Cliente |
| Índice | 0 | 1 | 2 | 3 | 4 | 5 |

Detección formato: cabecera[0] contiene "albar" → antigua (8 cols). Si no → nueva (6 cols).

### STOCK (fila 1 = título, fila 2 = cabecera, datos desde fila 3)
| Col | A | B | C | D | E | F |
|-----|---|---|---|---|---|---|
| Campo | Cliente | Referencia | Piezas/palet | Stock Palets | Stock Piezas | Estado |
| Índice | 0 | 1 | 2 | 3 | 4 | 5 |

### STOCK BASE (datos desde fila 3, 5 cols)
| Col | A | B | C | D | E |
|-----|---|---|---|---|---|
| Campo | Cliente | Referencia | Piezas/palet | Palets | Fecha |

Clave del mapa de stock: `cliente + '||' + referencia + '||' + piezasPal`

## Ficheros
| Fichero | Responsabilidad |
|---------|----------------|
| `stock.gs` | Núcleo: `actualizarStock`, `_leerFilasEntradas`, `_leerFilasSalidas`, `_rebuildEntradas`, `_rebuildSalidas`, `_leerStockBase`, `_consolidarFilas`, `_appendToMovimientos`, helpers de fecha/A1 |
| `setup.gs` | `getSS()` con el ID fijo, menú, `_setupEntradas/Salidas/Stock` |
| `export.gs` | `exportarInformeProveedor` → `_exportStock` (solo positivos, sin Estado). Nombre: `Informe_Almacen_JulioNunez_Duenas_YYYYMMDD.xlsx` |
| `webapp.gs` | `doGet`, `procesarMovimiento` — recibe datos del formulario |
| `index.html` | Formulario PWA móvil |

## Puntos clave de implementación
- `_consolidarFilas(rows, keyFn)`: agrupa por clave sumando palets+totPiezas. Usar siempre antes de escribir a la hoja.
- `_mergeGroups`: hace mergeVertically individualmente por celda, luego batch alignment con getRangeList — **RangeList NO soporta mergeVertically()**.
- `_aplicarBordes`: bordes externos + separadores de grupo con SOLID_MEDIUM.
- `actualizarStock(ss, sincronizar)`: si `sincronizar=true` hace rebuild completo de ENTRADAS y SALIDAS; si `false` solo recalcula STOCK (más rápido, para el formulario).
- Export elimina la columna Estado antes de guardar el xlsx.
