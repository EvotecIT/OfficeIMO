#region QuickStart
using System.Globalization;
using OfficeIMO.CSV;

CsvDocument.Create()
    .WithDelimiter(';')
    .WithCulture(CultureInfo.InvariantCulture)
    .WithHeader("Name", "Age", "City")
    .AddRow("Adele", 36, "Krakow")
    .AddRow("Lee", 30, "Krakow")
    .AddColumn("Bucket", row => row.AsInt32("Age") >= 35 ? "Senior" : "Regular")
    .SortBy("Age")
    .Save("People.csv", new CsvSaveOptions {
        Delimiter = ';',
        FormulaInjectionPolicy = CsvFormulaInjectionPolicy.Escape
    });
#endregion
