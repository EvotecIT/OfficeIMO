# Office-produced VBA fixtures

These small projects contain synthetic test source and no user document data.
Microsoft Office 16.0 on Windows produced the persisted module, directory, and
cache records. Tests use the files without requiring Office.

| Fixture | Producer and construction | SHA-256 |
| --- | --- | --- |
| `Excel-authored.bin` | Excel for Microsoft 365, build 20430. OfficeIMO created a workbook with a standard `Helpers.NativeSum` function returning 42. Excel opened and saved it, adding its document classes and compilation state. The native Visual Basic Editor inserted an empty ordinary `Class1`, then Excel saved the project. | `7d7b3a76a2181199721ff617e3d51a0b1b552be0222fe4f87609a0ac23fc914b` |
| `Word-authored.bin` | Word 16.0 created a blank macro-enabled document. The native Visual Basic Editor inserted an empty `Module1` in that document's project, then Word saved it. `ThisDocument` retains Word's `1Normal.ThisDocument` identity. | `dcb9f970078b3c0447d3f5fc81bed03df27fb8411f15228f470f6a243e90c418` |

The producer locale uses code page 1250. The tests register the encoding provider
already present in the test dependencies. Both projects are unprotected and
unsigned. The Excel fixture independently covers Office serialization and native
class authoring; its initial standard-module statements were authored by OfficeIMO.

Relevant contracts include logical names, class/document attributes, compatibility
version records, exact unchanged bytes, opaque compilation metadata, and source
edits that retain untouched module bytes. Whole-file signing and native source
loading are separate integration checks.
