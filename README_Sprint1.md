# Picking Orchestrator v4.0 — Sprint 1

App Streamlit unificada para automatización de picking — Beccacece Hnos SA.
Sucesora de `app_v3_9_1.py` (que **sigue funcionando intacta** como fallback).

## ✅ Qué incluye Sprint 1

| Tab | Estado | Notas |
|---|---|---|
| 1 — 📦 Planilla de Carga | ✅ Funcional | Reuso 100% de `generate_pdf` v3.9. Validado E2E con archivos reales: 76 págs / 15 repartos / 1.4s |
| 2 — 📋 Resumen Camiones | ✅ Funcional | Reuso 100% de `build_resumen_carga_pdf` v3.9. Validado: 2 págs / 10 camiones / 2.2s |
| 3 — 🚛 Camiones T2 | 🔧 Placeholder | Vista previa de datos disponible. Construcción real en Sprint 3 |
| 4 — 📊 Proyección Picking ×4 | 🔧 Placeholder | Sprint 4 |
| 5 — 📤 Extraíbles Sheets | ✅ Funcional | 4 bloques (Matriz Pall., Agregados AE, Reposición AE filtrada J<0, fx Picking) × 3 formatos (TSV/CSV/XLSX) |
| 6 — ✅ Validación + Log | 🔧 Placeholder | Sprint 2. El log de la corrida actual ya funciona |

**Sidebar:**
- 🧪 Toggle Dry-Run (procesa pero no genera PDFs finales)
- 💾 Toggle Snapshot (default ON, guarda en `./snapshots/YYYY-MM-DD/<run_id>/`)
- Stack técnico (versiones de libs)

## 📦 Instalación

```bash
# Python 3.11 recomendado (en Windows: poetry/venv si usás aislamiento)
pip install streamlit pandas openpyxl reportlab loguru
```

## ▶ Ejecución

```bash
streamlit run app_v4_0.py
```

Se abre en `http://localhost:8501`. La app es **autocontenida**: 1 solo `.py` de
1932 líneas, sin dependencias de la v3.9.

## 🧪 Cómo validar (recomendado)

1. Activar **Dry-Run** en sidebar.
2. Subir CAR.xlsx + Frescura 3.0.xlsx en Tab 1 → verificar preview.
3. Subir MASTER del día en Tab 5 → bajar Matriz Pall. en TSV → pegar en Google Sheets y comparar contra tu copy-paste manual.
4. Si todo coincide, desactivar Dry-Run y generar Planilla/Resumen reales en Tabs 1 y 2.
5. Verificar que los PDFs sean idénticos a los que generaba la v3.9 (deben serlo: reuso literal).

## 📋 Observaciones que detecté durante la inspección

1. **Matriz Picking APP — cols C-H**: el handoff (línea 132) decía `Cancha I/II/III/IV/MKPL/Comodín 1`, pero la realidad del MASTER es `CANCHA I/II/III/IV/MKPL/Comodín 1` — falta **CANCHA V** en esta hoja específica. En Matriz Pall. sí está CANCHA V. Asumo que el flujo de picking activo usa 5 canchas (I-IV + MKPL) y la V solo aplica a puras/AE. Confirmar cuando sea oportuno.
2. **Proyección Picking F6 col G**: hay una segunda fecha (6/5/2026) además de la principal (26/5/2026 en col C). Probable día comparativo histórico. A entender en Sprint 4.
3. **`load_frescura` devuelve 4 elementos** `(api, ddm, fr, diag)`, no 3 como decía el handoff (sección 12). Uso la firma real.
4. **MASTER de muestra**: 15 camiones con `Reparto=SI` de 29 totales. La mayoría con 0 pallets — para el día real, el filtro Tab 3 será `Reparto=SI ∧ TOTAL>0` típicamente ~17.

## 🛠 Próximos pasos (Sprint 2)

Construir Tab 6 (Validación + Reglas):
- Pre-checks: CAR/ANR modificados hoy, fecha A1 == hoy, suma TOTAL > 0
- Reglas: ocultar camiones TOTAL=0, tabla checkbox Reparto=NO
- Cross-validate CAR ↔ MASTER (warning, no bloqueante)
- Sanity checks bloqueantes antes de publicar
- Log de la corrida con timestamp y archivos procesados

## 🧩 Arquitectura

```
app_v4_0.py
├── Header v4.0 (líneas 1-65)
│   └── Imports, APP_VERSION, COLORS, setup_logger, log_event
├── Bloque reusado de v3.9 (líneas 66-1236)  ⚠ NO MODIFICAR
│   ├── Constantes (PAGE_W, DARK_BLUE, LEMAS, EXCLUDED_SKUS...)
│   ├── Utilidades (_norm, find_col, to_int_sku, normalize_cancha...)
│   ├── load_car, load_frescura, load_agr
│   ├── process_reparto, compute_pall_value
│   ├── draw_* (helpers de PDF — 30+ funciones)
│   ├── generate_pdf (Planilla)
│   └── build_resumen_carga_pdf (Resumen)
└── Extensiones v4.0 (líneas 1237-1932)
    ├── Lectores MASTER (Matriz Pall., Reposición AE, Picking APP, Agregados AE)
    ├── Lector T2 Pall. Camiones (referencia estática)
    ├── Exportadores (TSV/CSV/XLSX)
    ├── filter_reposicion_negativa
    ├── snapshot_local + make_run_id
    └── UI: render_sidebar, render_tab_*, main
```

**Decisión de diseño respetada**: 1 solo archivo autocontenido, copia literal del
core v3.9 (decisión cerrada Lucas). La v3.9 sigue intacta.
