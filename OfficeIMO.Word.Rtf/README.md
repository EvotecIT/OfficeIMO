# OfficeIMO.Word.Rtf

`OfficeIMO.Word.Rtf` bridges `OfficeIMO.Word` documents and the dependency-free `OfficeIMO.Rtf` engine.

```csharp
using OfficeIMO.Word;
using OfficeIMO.Word.Rtf;

using WordDocument word = WordDocument.Create();
word.AddParagraph("Hello RTF").SetBold();

string rtf = word.ToRtf();
WordDocument copy = rtf.LoadFromRtf();
```

The RTF parser and writer live in `OfficeIMO.Rtf`; this package only maps Word document objects to and from that reusable model.
