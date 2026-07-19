# OfficeIMO.Reader.Excel

Excel workbook ingestion for `OfficeIMO.Reader.Core`. Install this package when a reader needs XLSX, XLSM, XLSB, or XLS support, then compose it explicitly:

```csharp
using OfficeIMO.Reader;
using OfficeIMO.Reader.Excel;

OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
    .AddExcelHandler()
    .Build();
```
