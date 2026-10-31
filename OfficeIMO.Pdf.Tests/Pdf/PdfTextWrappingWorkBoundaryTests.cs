using System;
using System.Reflection;
using System.Threading;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfTextWrappingWorkBoundaryTests {
    [Fact]
    public void LineEndSplitBoundsMeasuredPrefixesForRepeatedZeroWidthBreaks() {
        string token = "MMMM" + string.Concat(System.Linq.Enumerable.Repeat("\u200B\u200C", 4000));
        long measuredCharacters = 0;
        Func<string, double> measure = text => {
            measuredCharacters += text.Length;
            return 40D;
        };
        MethodInfo split = typeof(PdfWriter).GetMethod("TrySplitTokenAtLineEnd",
            BindingFlags.NonPublic | BindingFlags.Static)!;
        object?[] arguments = { token, 36D, null, measure, null, 0D, 0, CancellationToken.None };

        Assert.False((bool)split.Invoke(null, arguments)!);
        Assert.InRange(measuredCharacters, 1L, Math.Max(4096L, 32L * token.Length));
    }

    [Fact]
    public void LineEndSplitObservesCancellationBetweenCandidateMeasurements() {
        string token = "MMMM" + string.Concat(System.Linq.Enumerable.Repeat("\u200B\u200C", 100));
        using var cancellation = new CancellationTokenSource();
        Func<string, double> measure = _ => {
            cancellation.Cancel();
            return 40D;
        };
        MethodInfo split = typeof(PdfWriter).GetMethod("TrySplitTokenAtLineEnd",
            BindingFlags.NonPublic | BindingFlags.Static)!;
        object?[] arguments = { token, 36D, null, measure, null, 0D, 0, cancellation.Token };

        TargetInvocationException error = Assert.Throws<TargetInvocationException>(() => split.Invoke(null, arguments));
        Assert.IsType<OperationCanceledException>(error.InnerException);
    }

    [Fact]
    public void LongCondensedTokenStillProducesReadablePdf() {
        var options = new PdfCreateOptions {
            DefaultFont = PdfStandardFont.TimesRoman,
            DefaultFontSize = OfficeLength.Points(10D),
            PageWidth = OfficeLength.Points(612D),
            PageHeight = OfficeLength.Points(792D),
            MarginLeft = OfficeLength.Points(36D),
            MarginRight = OfficeLength.Points(36D),
            MarginTop = OfficeLength.Points(36D),
            MarginBottom = OfficeLength.Points(36D)
        };
        var run = new PdfTextRun(new string('0', 4000) + new string('W', 200), fontSize: OfficeLength.Points(10D))
            .WithCharacterSpacing(OfficeLength.Points(-5D));

        byte[] pdf = PdfDocument.Create(options).Paragraph(paragraph => paragraph.Runs(new[] { run })).ToBytes();

        Assert.StartsWith("%PDF-", System.Text.Encoding.ASCII.GetString(pdf, 0, 5), StringComparison.Ordinal);
        Assert.Contains(new string('W', 20), PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }
}
