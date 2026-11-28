# XLSB migration verification

Run from the OfficeIMO checkout. The comparison uses separate worktrees for both
revisions. It copies tracked XLSB fixtures from the baseline and captures the
workbooks that the baseline XLSB tests save and reopen. It applies the same
operations to both revisions, compares fresh scalar and styled XLSB workbooks
from both writers, and writes a per-workbook report to
`comparison.json`. Supply `--corpus` to repeat the comparison with the same
inputs.

```sh
python3 Build/Verification/ExcelXlsb/compare.py origin/master HEAD "$comparison_dir" \
  --build-lock "$HOME/Developer/Scratch/coordinator/officeimo-build.lock"
```

Unchanged native saves must be byte-identical. For edited XLSB and XLSX export,
the comparison accepts ZIP container differences and compares the members.
Binary parts must match byte for byte. XML comparison treats namespace prefixes,
attribute order, and the order of relationship and content-type declarations
as equivalent. It resolves internal relationship targets and compares typed
XML relationship references by their target. For XLSX it can pair renamed parts
through the same relationship edges and compares their content types and bytes
under that mapping. An omitted worksheet dimension is equivalent to one that
exactly bounds the stored cells or the stored cells and their merged ranges.
The cells and merges themselves must match. It ignores the modification time written on save and compares the remaining core properties. Refusals must have
the same exception type and message.

Run the paired benchmark on the Mac mini. The lock covers both builds and all
measurements. The baseline executable generates the plain input workbook and
the workbook with array formulas. Each revision validates the cell values and
array metadata. The runner records raw time and allocation samples with
alternating revision order.

```sh
python3 Build/Verification/ExcelXlsb/benchmark.py origin/master HEAD "$benchmark_dir" \
  --lock /Volumes/Samsung/Development/Scratch/officeimo-bench.lock
```
