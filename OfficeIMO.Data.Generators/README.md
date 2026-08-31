# OfficeIMO.Data.Generators

`OfficeIMO.Data.Generators` creates allocation-light, reflection-free configuration for
`OfficeIMO.Data.RowMapper<T>`.

Add the generator beside the OfficeIMO package that supplies the reader:

```powershell
dotnet add package OfficeIMO.CSV
dotnet add package OfficeIMO.Data.Generators
```

```csharp
using OfficeIMO.Data;

[GenerateRowMapper]
public sealed class InvoiceRow {
    [DataColumn("Invoice Id", "Id")]
    public int InvoiceId { get; set; }

    public decimal Total { get; set; }
}

IEnumerable<InvoiceRow> rows = reader.RowsAs<InvoiceRow>(InvoiceRowRowMapping.Configure);
```

The generated mapper is shared by Excel, CSV, and any other `DbDataReader`. No runtime
reflection or expression compilation is required. The annotated model must be a concrete,
top-level, non-generic class or struct with at least one writable public property; classes also
need a public parameterless constructor. Writable inherited properties are included, and the
most-derived declaration wins when a property is hidden. Ref-like, abstract, nested, generic,
and file-local shapes produce build diagnostics instead of a runtime fallback.
