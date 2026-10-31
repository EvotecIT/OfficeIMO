#region QuickStart
using OfficeIMO;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;

Sale[] sales = [new("North", 182_400m, 0.12), new("South", 141_900m, 0.08), new("West", 97_300m, -0.03)];

using (var report = WordDocument.Create("Sales.docx")) {
    report.AddHeading("Quarterly sales", level: 1);
    report.AddParagraph("Revenue by region for the third quarter.");
    report.AddTable(sales);
    report.AddChart(OfficeChartType.Column, sales, nameof(Sale.Region), nameof(Sale.Revenue));
    report.Save();
    report.SaveAsPdf("Sales.pdf");
}

using (var book = ExcelDocument.Create("Sales.xlsx", new ExcelCreateOptions { Theme = OfficeTheme.Modern })) {
    ExcelSheet sheet = book.AddSheet("Sales");
    sheet.AddTable(sales, new OfficeTableOptions { Name = "Sales" });
    sheet.AddChart(OfficeChartType.Column, sales, nameof(Sale.Region),
        new ExcelChartOptions { Title = "Revenue by region", Row = 2, Column = 6 }, nameof(Sale.Revenue));
    book.Save();
    book.SaveAsPdf("Sales-workbook.pdf");
}

using (var deck = PowerPointDocument.Create("Sales.pptx")) {
    deck.AddSlide().AddTitle("Quarterly sales");
    PowerPointSlide chartSlide = deck.AddSlide();
    chartSlide.AddTitle("Revenue by region");
    chartSlide.AddChart(OfficeChartType.Column, sales, nameof(Sale.Region), new PowerPointChartOptions {
        Left = OfficeLength.Centimeters(2.3), Top = OfficeLength.Centimeters(5.5),
        Width = OfficeLength.Centimeters(29), Height = OfficeLength.Centimeters(12)
    }, nameof(Sale.Revenue));
    deck.Save();
    deck.SaveAsPdf("Sales-deck.pdf");
}

record Sale(string Region, decimal Revenue, double Growth);
#endregion
