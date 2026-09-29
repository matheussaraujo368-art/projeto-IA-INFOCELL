# Dados brutos

Arquivos **sem nenhuma alteração** em relação à fonte original.

| Arquivo | Origem | Observação |
|---|---|---|
| `ora_202507_dispositivos_bruto.csv` | `aggregated/202507/OpenRepairData_v0.3_aggregate_202507.csv` | Apenas as linhas de `product_category` ∈ {Mobile, Laptop, Desktop computer, Tablet} (23.612 registros). Nenhuma coluna ou valor alterado. |
| `OpenRepairData_v0.3_Product_Categories.csv`, `tableschema.json` | `aggregated/202507/` | Categorias e esquema oficial ORDS v0.3 |
| `quests/mobifix.csv` | `quests/mobiles/` | 1.916 celulares com tipo de falha |
| `quests/compcat_laptops.csv`, `compcat_desktops.csv`, `compcat_fault_types.csv` | `quests/computers/` | 3.510 notebooks + 583 desktops |
| `quests/tabicat.csv`, `tabicat_fault_types.csv` | `quests/tablets/` | 647 tablets (com idioma rotulado) |

## Fonte e versão
- Open Repair Alliance — https://openrepair.org/open-data/
- Repositório: https://github.com/openrepair/data (commit `90eba80740506b4fa283288780d34e9e7e01deef`, 22/03/2026)
- Padrão ORDS v0.3, exportação `202507` (dados até 31/07/2025; 305.649 registros no total)
- Licença: **CC BY-SA 4.0** — os dados tratados deste repositório seguem a mesma licença.

## Base completa
O CSV agregado completo (58 MB) não é versionado. Para obtê-lo:
```bash
python data/raw/baixar_base_completa.py
```
Se `ora_202507_dispositivos_bruto.csv` for apagado, `src/preparacao.py` o recria a partir do arquivo completo.
