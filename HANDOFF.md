# HANDOFF.md

## Current state

The BOM converter is implemented as a PyQt5 Windows GUI and runs on the current workstation through `BOM_변환기/실행.bat`.

Users can drag multiple numeric `.xls`/`.xlsx` RAW files or a folder into a queue, select a blank template and output folder, and convert sequentially. The GUI reports each file and overall progress, continues after individual errors, supports cancellation of waiting jobs, protects existing outputs unless overwrite is selected, and writes a CSV log.

The supplied empty workbook is bundled at `BOM_변환기/templates/BOM 자동 입력 양식.xlsx`. Outputs are named `<제품코드> 완료.xlsx`.

## Implemented mapping

| Output | Source/default |
| --- | --- |
| A `ITEM_CD` | RAW S, parent item code |
| B `M_UNIT_CD` | `KG` |
| C `PRPOS_CD` | `M` |
| D `ALT_NO` | `0` |
| E `BOM_QT` | `1` |
| F `MTRIL_CD` | RAW W, child item code |
| G `UNIT_CD` | RAW Z exactly; no unit conversion |
| H `NET_QT` | RAW AA, required quantity |
| I `SCRAP_RT` | Direct child count of the child item |
| J:T | Blank |

Rows are sorted by BOM level and source-row order. The converter validates headers, levels, filename/product consistency, codes, quantity range/precision, and non-empty units.

## Important files

- `BOM_변환기/app.py`: GUI, queue, worker thread, progress, cancellation, and logs.
- `BOM_변환기/bom_core.py`: parsers, mapping, template preservation, verification, and atomic output.
- `BOM_변환기/artifact_engine.mjs`: Artifact Tool workbook creation.
- `BOM_변환기/실행.bat`: Windows launcher.
- `BOM_변환기/README.md`: end-user instructions.
- `BOM_변환기/tests/test_package.py`: integration/regression suite.

## Verification completed

The integration suite passed ten scenarios:

1. Qt file drop and duplicate-path suppression.
2. Sequential GUI conversion of `110585.xls` and `150149.xls`.
3. Independent verification of 51 BOM rows / 1,020 output data cells and template styles.
4. Existing-file protection.
5. A 1,200-file folder queue.
6. A 57-row BOM extending beyond the 32-row blank template.
7. Invalid input followed by successful processing of the next file.
8. Duplicate product codes in different directories.
9. Cancellation after the current file.
10. Exact unit preservation and rejection of non-finite quantities.

The BAT/Qt offscreen smoke test passed. Evidence is in `BOM_변환기/tests/test_results.json`; generated workbooks and logs are ignored.

## Known constraints and open decisions

- The BAT depends on this workstation's Codex bundled Python, Node, and Artifact Tool runtime. PyQt5 and `xlrd` are vendored, but this is not a standalone EXE.
- Column I is a direct-child count because that exactly reproduced the handwritten example. The template describes it as Scrap rate, so its final ERP meaning must be confirmed.
- `ALT_NO=0` follows the test example although the template guidance mentions `01`–`99`; confirm the accepted ERP value.
- B is fixed to `KG`. G preserves RAW Z exactly and performs no conversion.
- Files are independent. Shared subassembly BOMs across files are not merged or deduplicated.
- Binary `.xls` and `.xlsx` parsing is implemented, while the supplied production samples were CP949 tab-delimited data with an `.xls` suffix.

## Recommended next work

Create a feature branch from `main` before new development, preferably `codex/<feature>`.

1. Confirm Scrap and alternative-number rules and add fixtures for the final behavior.
2. Validate several hundred representative RAW files and classify exceptions.
3. Add a conflict report for repeated subassembly BOMs across products.
4. If deployment without Codex is needed, package a standalone signed installer/EXE or replace the Artifact Tool dependency.
