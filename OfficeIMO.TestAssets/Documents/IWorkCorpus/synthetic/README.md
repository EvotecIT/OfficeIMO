# Synthetic iWork controls

`numbers-unpadded-scalar.numbers` contains one sheet, one table and the numeric value 42. It has no cell-padding declaration. It exercises the complete Numbers-to-PDF route separately from Apple-produced tables whose padding currently prevents complete conversion.

This is a format-boundary control under the repository's MIT license. It does not establish Apple application acceptance, appearance or general Numbers fidelity. The workflow correctness suite and NativeAOT host use the same bytes.

Reproduce the package from the repository root:

```sh
python3 Build/IWork/generate-unpadded-pdf-control.py OfficeIMO.TestAssets/Documents/IWorkCorpus/synthetic/numbers-unpadded-scalar.numbers
```
