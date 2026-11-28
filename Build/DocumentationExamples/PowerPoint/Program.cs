#region QuickStart
using OfficeIMO;
using OfficeIMO.PowerPoint;

Region[] regions = [new("North", 5.7m, 0.36), new("South", 4.1m, 0.31), new("West", 2.9m, 0.27)];

using var deck = PowerPointDocument.Create("Regions.pptx");
PowerPointLayoutBox page = deck.SlideSize.GetContentBox(OfficeLength.Centimeters(1.5));
PowerPointLayoutBox titleBand = page.TakeTop(OfficeLength.Centimeters(2.5));
PowerPointLayoutBox[] halves = page.Inset(OfficeLength.Zero, OfficeLength.Centimeters(3.5), OfficeLength.Zero, OfficeLength.Zero)
    .SplitColumns(2, OfficeLength.Centimeters(1));

PowerPointSlide slide = deck.AddSlide();
slide.AddTitle("Revenue by region", titleBand);
halves[0].ApplyTo(slide.AddTable(regions));
slide.AddChart(OfficeChartType.Column, regions, nameof(Region.Name), new PowerPointChartOptions {
    Title = "Revenue (USD m)",
    Left = halves[1].Left, Top = halves[1].Top, Width = halves[1].Width, Height = halves[1].Height
}, nameof(Region.Revenue));

deck.Save();

record Region(string Name, decimal Revenue, double Margin);
#endregion
