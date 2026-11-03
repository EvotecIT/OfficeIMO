using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfLayoutDiagnosticOwnershipTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void NestedConversionCompletesItsOwnLossReportBeforeRunningCallbackReturns(bool abortOuter) {
        var controlReport = new PdfConversionReport();
        byte[] control = WriteClippedDocument(controlReport);
        AssertClippingWarning(controlReport.Warnings);
        Assert.True(controlReport.HasLoss);
        using (var pdf = PdfPigDocument.Open(control)) Assert.Contains("First", pdf.GetPage(1).Text);

        var outerReport = new PdfConversionReport();
        var observations = new List<InnerObservation>();
        var outer = PdfDocument.Create(Options().ReportDiagnosticsTo(outerReport, "Outer"));
        outer.Header(header => header.Content(_ => content => {
            var innerReport = new PdfConversionReport();
            byte[] inner = WriteClippedDocument(innerReport);
            observations.Add(new InnerObservation(innerReport, innerReport.HasLoss, innerReport.Warnings.ToArray()));
            SaveEvidence(abortOuter ? "inner-before-aborted-outer" : "inner-before-completed-outer", inner);
            if (abortOuter) throw new OuterProbeAbortedException();
            content.AddParagraph(paragraph => paragraph.AddText("HEADER"));
        }, global::OfficeIMO.OfficeLength.Points(18D)));
        outer.Paragraph(paragraph => paragraph.AddText("OUTER"));

        if (abortOuter) Assert.Throws<OuterProbeAbortedException>(() => outer.ToBytes());
        else {
            using var pdf = PdfPigDocument.Open(outer.ToBytes());
            Assert.Contains("OUTER", pdf.GetPage(1).Text);
            Assert.Contains("HEADER", pdf.GetPage(1).Text);
        }

        Assert.NotEmpty(observations);
        Assert.All(observations, observation => {
            Assert.True(observation.HasLossAtReturn, "The completed inner conversion must report its loss before its caller returns.");
            AssertClippingWarning(observation.WarningsAtReturn);
            Assert.True(observation.Report.HasLoss);
            AssertClippingWarning(observation.Report.Warnings);
        });
        Assert.Empty(outerReport.Warnings);
        Assert.False(outerReport.HasLoss);
    }

    private static byte[] WriteClippedDocument(PdfConversionReport report) =>
        PdfDocument.Create(Options().ReportDiagnosticsTo(report, "Inner"))
            .Canvas(canvas => canvas.TextBox("First\nSecond\nThird", global::OfficeIMO.OfficeLength.Points(24D), global::OfficeIMO.OfficeLength.Points(24D), global::OfficeIMO.OfficeLength.Points(120D), global::OfficeIMO.OfficeLength.Points(18D),
                new PdfCanvasTextBoxStyle {
                    FontSize = global::OfficeIMO.OfficeLength.Points(9D), PaddingX = global::OfficeIMO.OfficeLength.Points(0D), PaddingY = global::OfficeIMO.OfficeLength.Points(0D), Background = null, BorderColor = null
                })).ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });

    private static void AssertClippingWarning(IReadOnlyList<PdfConversionWarning> warnings) {
        PdfConversionWarning warning = Assert.Single(warnings, item => item.Code == "canvas-text-box-content-clipped");
        Assert.Equal("Inner", warning.Converter);
        Assert.Equal(PdfLayoutDiagnosticKind.ClippedContent, warning.LayoutDiagnostic!.Kind);
        Assert.Equal(18D, (warning.LayoutDiagnostic.Height?.ToPoints()));
    }

    private static PdfCreateOptions Options() => new() {
        PageSetup = {
            Size = new PdfPageSize(global::OfficeIMO.OfficeLength.Points(240D), global::OfficeIMO.OfficeLength.Points(320D)),
            Margins = { Left = global::OfficeIMO.OfficeLength.Points(24D), Right = global::OfficeIMO.OfficeLength.Points(24D), Top = global::OfficeIMO.OfficeLength.Points(32D), Bottom = global::OfficeIMO.OfficeLength.Points(32D) },
            Header = { Show = false },
            Footer = { Show = false }
        }, DefaultFont = PdfStandardFont.Helvetica,
        DefaultFontSize = global::OfficeIMO.OfficeLength.Points(11D)
    };

    private static void SaveEvidence(string name, byte[] bytes) {
        string directory = Path.Combine(AppContext.BaseDirectory, "diagnostic-scope-artifacts");
        Directory.CreateDirectory(directory);
        File.WriteAllBytes(Path.Combine(directory, name + ".pdf"), bytes);
    }

    private sealed record InnerObservation(PdfConversionReport Report, bool HasLossAtReturn,
        PdfConversionWarning[] WarningsAtReturn);
    private sealed class OuterProbeAbortedException : Exception { }
}
