## Mail merge

OfficeIMO.Word can replace `MERGEFIELD` entries with provided values. Fields can either be converted to plain text or kept for later updates.

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.AddParagraph("Dear ")
        .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"FirstName\"" })
        .AddText(",");

    document.AddParagraph("Your order number ")
        .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"OrderId\"" })
        .AddText(" has been processed.");

    var values = new Dictionary<string, string> {
        { "FirstName", "John" },
        { "OrderId", "12345" }
    };

    WordMailMerge.Execute(document, values);
    document.Save(true);
}
```

```csharp
using (WordDocument document = WordDocument.Create(filePath)) {
    document.AddParagraph("Hello ")
        .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"FirstName\"" })
        .AddText(" ")
        .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"LastName\"" })
        .AddText(",");

    document.AddParagraph("Your balance is ")
        .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"Balance\"" })
        .AddText(".");

    var values = new Dictionary<string, string> {
        { "FirstName", "Jane" },
        { "LastName", "Doe" },
        { "Balance", "$200" }
    };

    WordMailMerge.Execute(document, values, removeFields: false);
    document.Save(true);
}
```
