using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Word;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_ChartSeriesNamesAreLiteralWithoutWorkbookReferences() {
            string filePath = Path.Combine(_directoryWithFiles, "ChartSeriesNamesLiteral.docx");
            var categories = new List<string> { "Food", "Housing", "Mix" };
            using (WordDocument document = WordDocument.Create(filePath)) {
                var bar = document.AddChart("Bar");
                bar.AddCategories(categories);
                bar.AddBar("Brazil", new[] { 10, 35, 18 }, OfficeColor.Brown);
                var line = document.AddChart("Line");
                line.AddChartAxisX(categories);
                line.AddLine("Poland", new List<int> { 13, 20, 23 }, OfficeColor.Green);
                var area = document.AddChart("Area");
                area.AddChartAxisX(categories);
                area.AddArea("USA", new List<int> { 3, 5, 8 }, OfficeColor.Blue);
                var scatter = document.AddChart("Scatter");
                scatter.AddScatter("Points", new List<double> { 1, 2, 3 }, new List<double> { 4, 5, 6 }, OfficeColor.Red);
                var radar = document.AddChart("Radar");
                radar.AddChartAxisX(categories);
                radar.AddRadar("Radar", new List<int> { 1, 2, 3 }, OfficeColor.Orange);
                document.Save();
            }

            using WordprocessingDocument package = WordprocessingDocument.Open(filePath, false);
            var chartParts = package.MainDocumentPart!.ChartParts.ToList();
            Assert.Equal(5, chartParts.Count);
            var names = new List<string>();
            foreach (ChartPart part in chartParts) {
                foreach (C.SeriesText text in part.ChartSpace!.Descendants<C.SeriesText>()) {
                    // Without an embedded workbook there is no cell to reference, so the name is a literal.
                    Assert.Empty(text.Descendants<C.StringReference>());
                    names.Add(Assert.Single(text.Elements<C.NumericValue>()).Text);
                }
            }
            Assert.Equal(new[] { "Brazil", "Poland", "USA", "Points", "Radar" }, names);
        }
    }
}
