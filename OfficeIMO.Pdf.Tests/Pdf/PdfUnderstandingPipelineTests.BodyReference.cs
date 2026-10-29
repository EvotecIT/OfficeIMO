using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfUnderstandingPipelineTests {
    [Fact]
    public void BodyReferenceSize_ComesFromRunningTextRatherThanChartLabels() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var options = PdfUnderstandingPipelineOptions.Structured();
        // A dashboard: 11 pt prose beside a chart whose 9 pt axis and legend labels outnumber the prose lines.
        options.GlyphDecoding = new FixedGlyphStage(new[] {
            new PdfTextSpan("Quarterly operations dashboard", "F1", 20D, 40D, 760D, 300D),
            new PdfTextSpan("Service levels, resolved volume and the decisions needed for Q4.", "F1", 11D, 40D, 735D, 320D),
            new PdfTextSpan("Resolved", "F1", 9D, 40D, 690D, 40D),
            new PdfTextSpan("Target", "F1", 9D, 90D, 690D, 30D),
            new PdfTextSpan("July", "F1", 9D, 40D, 560D, 20D),
            new PdfTextSpan("August", "F1", 9D, 100D, 560D, 30D),
            new PdfTextSpan("September", "F1", 9D, 160D, 560D, 45D),
            new PdfTextSpan("Region", "F1", 9D, 40D, 520D, 30D),
            new PdfTextSpan("Owner", "F1", 9D, 40D, 500D, 30D),
            new PdfTextSpan("SLA attainment finished above the target in all three months", "F1", 11D, 330D, 690D, 260D),
            new PdfTextSpan("while resolved volume grew. Integrations is the only service", "F1", 11D, 330D, 675D, 260D),
            new PdfTextSpan("line below target, and its response time is the main driver.", "F1", 11D, 330D, 660D, 260D)
        });

        PdfLogicalPage page = Assert.Single(Read(pdf, options).Pages);

        Assert.Contains(page.Headings, heading => heading.Text == "Quarterly operations dashboard");
        Assert.DoesNotContain(page.Headings, heading => heading.Text.StartsWith("Service levels", StringComparison.Ordinal));
        Assert.DoesNotContain(page.Headings, heading => heading.Text.IndexOf("SLA attainment", StringComparison.Ordinal) >= 0);
        Assert.DoesNotContain(page.Headings, heading => heading.Text.IndexOf("main driver", StringComparison.Ordinal) >= 0);
    }

    [Fact]
    public void NumberedLineInHeadingSize_IsHeadingWhileBodySizedStepsStayListItems() {
        byte[] pdf = PdfDocument.Create().Paragraph(p => p.Text("placeholder")).ToBytes();
        var options = PdfUnderstandingPipelineOptions.Structured();
        options.GlyphDecoding = new FixedGlyphStage(new[] {
            new PdfTextSpan("2. Revenue by region", "F1", 17D, 50D, 740D, 200D),
            new PdfTextSpan("Table 1 shows revenue by region for the first two quarters of the year.", "F1", 11D, 50D, 712D, 400D),
            new PdfTextSpan("Figures are unaudited and stated in US dollars throughout this report.", "F1", 11D, 50D, 698D, 400D),
            new PdfTextSpan("4. Next steps", "F1", 17D, 50D, 660D, 130D),
            new PdfTextSpan("1. Approve the second-source budget for R-01.", "F1", 11D, 50D, 632D, 300D),
            new PdfTextSpan("2. Confirm the hedging policy change.", "F1", 11D, 50D, 616D, 260D)
        });

        PdfLogicalPage page = Assert.Single(Read(pdf, options).Pages);

        Assert.Contains(page.Headings, heading => heading.Text == "2. Revenue by region");
        Assert.Contains(page.Headings, heading => heading.Text == "4. Next steps");
        Assert.DoesNotContain(page.ListItems, item => item.Text.IndexOf("Next steps", StringComparison.Ordinal) >= 0);
        Assert.Contains(page.ListItems, item => item.Text.IndexOf("Approve the second-source budget", StringComparison.Ordinal) >= 0);
        Assert.Contains(page.ListItems, item => item.Text.IndexOf("Confirm the hedging policy change", StringComparison.Ordinal) >= 0);
    }
}
