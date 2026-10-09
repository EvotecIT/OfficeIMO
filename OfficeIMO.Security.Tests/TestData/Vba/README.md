# Office-produced VBA fixtures

These small projects contain synthetic test source and no user document data.
Microsoft Office 16.0 on Windows produced the two locally authored projects.
The independently supplied nested-form fixture is attributed below. Tests use
these persisted project records without requiring Office.

| Fixture | Producer and construction | SHA-256 |
| --- | --- | --- |
| `Excel-authored.bin` | Excel for Microsoft 365, build 20430. OfficeIMO created a workbook with a standard `Helpers.NativeSum` function returning 42. Excel opened and saved it, adding its document classes and compilation state. The native Visual Basic Editor inserted an empty ordinary `Class1`, then Excel saved the project. | `7d7b3a76a2181199721ff617e3d51a0b1b552be0222fe4f87609a0ac23fc914b` |
| `Word-authored.bin` | Word 16.0 created a blank macro-enabled document. The native Visual Basic Editor inserted an empty `Module1` in that document's project, then Word saved it. `ThisDocument` retains Word's `1Normal.ThisDocument` identity. | `dcb9f970078b3c0447d3f5fc81bed03df27fb8411f15228f470f6a243e90c418` |
| `Excel-nested-form.bin` | The unmodified `xl/vbaProject.bin` extracted from an independently supplied Excel workbook containing `FrmNested` and nested form controls. Tests edit form source and verify all 17 designer streams and storage metadata. | `533bb3f8b75067bb34901e3b39a2b63d0583edcee93833075d1056e8e87bf9da` |

The first two fixtures use code page 1250. The tests register the encoding provider
already present in the test dependencies. Both projects are unprotected and
unsigned. The Excel fixture independently covers Office serialization and native
class authoring; its initial standard-module statements were authored by OfficeIMO.

The nested-form fixture uses code page 1252. Its source is
[pyOpenVBA's Excel-produced test fixture](https://github.com/WilliamSmithEdward/pyOpenVBA/blob/2b41b07ca817538be628cd4591d13a9b2c92c058/tests/live_excel_testing/nested_form.xlsm).
The workbook SHA-256 is `efacfe427762c94cafec645b2931f25b0c4ab15e9c55d0b96de2e7051c5a6ba6`.
It is included only as test data under its [MIT license](Excel-nested-form.LICENSE.md);
no external library or runtime is used by these tests.

Relevant contracts include logical names, class/document attributes, compatibility
version records, exact unchanged bytes, opaque compilation metadata, and source
edits that retain untouched module bytes. Whole-file signing and native source
loading are separate integration checks.
