using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfLogicalTableValueAnalysisTests {
    [Fact]
    public void Extract_ProfilesRichTableValueKindsForAllAdapters() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
                PageWidth = 760,
                PageHeight = 360,
                MarginLeft = 20,
                MarginRight = 20,
                MarginTop = 20,
                MarginBottom = 20,
                DefaultFontSize = 9
            })
            .Table(new[] {
                new[] { "Active", "Completion", "Due Date", "Start Time", "Quantity", "Owner" },
                new[] { "True", "25%", "2026-08-18", "09:30", "12.5", "Ada" },
                new[] { "False", "100%", "2026-08-19", "17:45", "3", "Grace" }
            }, style: new PdfTableStyle {
                HeaderRowCount = 1,
                ColumnWidthPoints = new List<double?> { 80, 100, 110, 100, 90, 100 }
            })
            .ToBytes();

        PdfLogicalTable table = Assert.Single(Assert.Single(PdfDocumentReadResult.Load(pdf).Pages).Tables);
        PdfLogicalTableData data = PdfLogicalTableAnalysis.Extract(table);

        Assert.Equal(
            new[] {
                PdfLogicalTableValueKind.Boolean,
                PdfLogicalTableValueKind.Percentage,
                PdfLogicalTableValueKind.DateTime,
                PdfLogicalTableValueKind.Time,
                PdfLogicalTableValueKind.Number,
                PdfLogicalTableValueKind.Text
            },
            data.ValueProfiles.Select(static profile => profile.Kind));
        Assert.All(data.ValueProfiles, profile => Assert.Equal(1D, profile.Confidence));
        Assert.All(data.ValueProfiles, profile => Assert.Equal(2, profile.NonEmptyCellCount));
    }

    [Fact]
    public void Analyze_MixedTypedAndTextValuesReportEvidenceBasedConfidence() {
        IReadOnlyList<IReadOnlyList<string>> rows = new[] {
            (IReadOnlyList<string>) new[] { "1" },
            new[] { "2" },
            new[] { "N/A" }
        };

        PdfLogicalTableValueProfile profile = Assert.Single(
            PdfLogicalTableValueAnalysis.Analyze(new[] { "Quantity" }, rows));

        Assert.Equal(PdfLogicalTableValueKind.Text, profile.Kind);
        Assert.Equal(3, profile.NonEmptyCellCount);
        Assert.Equal(1, profile.MatchingCellCount);
        Assert.Equal(1D / 3D, profile.Confidence, 8);
    }

    [Theory]
    [InlineData("Yes", "No")]
    [InlineData("Ja", "Nej")]
    [InlineData("Да", "Нет")]
    [InlineData("是", "否")]
    public void Analyze_DoesNotUseNaturalLanguageWordsAsBooleanSyntax(string trueWord, string falseWord) {
        IReadOnlyList<IReadOnlyList<string>> rows = new[] {
            (IReadOnlyList<string>) new[] { trueWord },
            new[] { falseWord }
        };

        PdfLogicalTableValueProfile profile = Assert.Single(
            PdfLogicalTableValueAnalysis.Analyze(new[] { "状态" }, rows));

        Assert.Equal(PdfLogicalTableValueKind.Text, profile.Kind);
    }

    [Fact]
    public void Analyze_UsesLocalizedDatesOnlyWithAnExplicitCulture() {
        IReadOnlyList<IReadOnlyList<string>> rows = new[] {
            (IReadOnlyList<string>) new[] { "31.12.2026" },
            new[] { "01.01.2027" }
        };

        PdfLogicalTableValueProfile generic = Assert.Single(
            PdfLogicalTableValueAnalysis.Analyze(new[] { "Termin" }, rows));
        PdfLogicalTableValueProfile polish = Assert.Single(
            PdfLogicalTableValueAnalysis.Analyze(
                new[] { "Termin" },
                rows,
                System.Globalization.CultureInfo.GetCultureInfo("pl-PL")));

        Assert.Equal(PdfLogicalTableValueKind.Text, generic.Kind);
        Assert.Equal(PdfLogicalTableValueKind.DateTime, polish.Kind);
    }

    [Fact]
    public void Analyze_DoesNotMisclassifyIsoDateTimesAsClockTimes() {
        IReadOnlyList<IReadOnlyList<string>> rows = new[] {
            (IReadOnlyList<string>) new[] { "2026-08-18 09:30" },
            new[] { "2026-08-19 17:45" }
        };

        PdfLogicalTableValueProfile profile = Assert.Single(
            PdfLogicalTableValueAnalysis.Analyze(new[] { "Timestamp" }, rows));

        Assert.Equal(PdfLogicalTableValueKind.DateTime, profile.Kind);
    }

    [Theory]
    [InlineData("abc123")]
    [InlineData("12(3)")]
    [InlineData("123kg")]
    public void TryParseNumericValue_RejectsTextInsteadOfSalvagingEmbeddedDigits(string source) {
        Assert.False(PdfLogicalTableAnalysis.TryParseNumericValue(source, null, out _));
    }

    [Theory]
    [InlineData("１２３．５０", "123.50")]
    [InlineData("١٢٣٫٥٠", "123.50")]
    [InlineData("٠٫١٢٥", "0.125")]
    [InlineData("０．１２５", "0.125")]
    [InlineData("𝟘.𝟙𝟚𝟝", "0.125")]
    [InlineData("١٬٢٣٤", "1234")]
    [InlineData("𝟙𝟚.𝟝", "12.5")]
    [InlineData("（１，２３４．５）", "-1234.5")]
    public void TryParseNumericValue_NormalizesUnicodeDecimalSyntax(string source, string expected) {
        Assert.True(PdfLogicalTableAnalysis.TryParseNumericValue(source, null, out decimal actual));
        Assert.Equal(decimal.Parse(expected, System.Globalization.CultureInfo.InvariantCulture), actual);
    }

    [Fact]
    public void TryParseNumericValue_UsesCultureInsteadOfDigitCountForUnicodeSeparators() {
        Assert.True(PdfLogicalTableAnalysis.TryParseNumericValue("𝟙.𝟚𝟛𝟜", null, out decimal invariant));
        Assert.Equal(1.234m, invariant);

        Assert.True(PdfLogicalTableAnalysis.TryParseNumericValue(
            "𝟙,𝟚𝟛𝟜",
            System.Globalization.CultureInfo.GetCultureInfo("pl-PL"),
            out decimal polish));
        Assert.Equal(1.234m, polish);
    }

    [Fact]
    public void Analyze_RecognizesUnicodeDecimalDigitsAndPercentSigns() {
        IReadOnlyList<IReadOnlyList<string>> rows = new[] {
            (IReadOnlyList<string>) new[] { "٢٥٪" },
            new[] { "１００％" }
        };

        PdfLogicalTableValueProfile profile = Assert.Single(
            PdfLogicalTableValueAnalysis.Analyze(new[] { "نسبة" }, rows));

        Assert.Equal(PdfLogicalTableValueKind.Percentage, profile.Kind);
        Assert.Equal(1D, profile.Confidence);
    }
}
