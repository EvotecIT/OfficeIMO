using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfLogicalTableValueAnalysisTests {
    [Fact]
    public void ValueKind_AddsCurrencyWithoutRenumberingExistingMembers() {
        Assert.Equal(0, (int)PdfLogicalTableValueKind.Empty);
        Assert.Equal(1, (int)PdfLogicalTableValueKind.Text);
        Assert.Equal(2, (int)PdfLogicalTableValueKind.Number);
        Assert.Equal(3, (int)PdfLogicalTableValueKind.Percentage);
        Assert.Equal(4, (int)PdfLogicalTableValueKind.Boolean);
        Assert.Equal(5, (int)PdfLogicalTableValueKind.DateTime);
        Assert.Equal(6, (int)PdfLogicalTableValueKind.Time);
        Assert.Equal(7, (int)PdfLogicalTableValueKind.Currency);
    }

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
                new PdfLogicalTableValueAnalysisOptions {
                    DateTimeCulture = System.Globalization.CultureInfo.GetCultureInfo("pl-PL")
                }));

        Assert.Equal(PdfLogicalTableValueKind.Text, generic.Kind);
        Assert.Equal(PdfLogicalTableValueKind.DateTime, polish.Kind);
    }

    [Fact]
    public void Analyze_KeepsNumericAndDateTimeCulturePoliciesIndependent() {
        IReadOnlyList<IReadOnlyList<string>> rows = new[] {
            (IReadOnlyList<string>) new[] { "1 234,50", "01.02.2025" },
            new[] { "2 345,75", "03.04.2025" }
        };

        IReadOnlyList<PdfLogicalTableValueProfile> numericOnly = PdfLogicalTableValueAnalysis.Analyze(
            new[] { "Kwota", "Termin" },
            rows,
            new PdfLogicalTableValueAnalysisOptions {
                NumericCulture = System.Globalization.CultureInfo.GetCultureInfo("pl-PL")
            });
        IReadOnlyList<PdfLogicalTableValueProfile> localizedDates = PdfLogicalTableValueAnalysis.Analyze(
            new[] { "Kwota", "Termin" },
            rows,
            new PdfLogicalTableValueAnalysisOptions {
                NumericCulture = System.Globalization.CultureInfo.GetCultureInfo("pl-PL"),
                DateTimeCulture = System.Globalization.CultureInfo.GetCultureInfo("pl-PL")
            });

        Assert.Equal(PdfLogicalTableValueKind.Number, numericOnly[0].Kind);
        Assert.Equal(PdfLogicalTableValueKind.Text, numericOnly[1].Kind);
        Assert.Equal(PdfLogicalTableValueKind.DateTime, localizedDates[1].Kind);
    }

    [Fact]
    public void Analyze_DoesNotTreatLocalizedDottedDatesAsGroupedIntegers() {
        var culture = System.Globalization.CultureInfo.GetCultureInfo("de-DE");
        IReadOnlyList<IReadOnlyList<string>> rows = new[] {
            (IReadOnlyList<string>) new[] { "31.12.2026" },
            new[] { "30.11.2026" }
        };

        PdfLogicalTableValueProfile numericOnly = Assert.Single(
            PdfLogicalTableValueAnalysis.Analyze(
                new[] { "Termin" },
                rows,
                new PdfLogicalTableValueAnalysisOptions { NumericCulture = culture }));
        PdfLogicalTableValueProfile localizedDates = Assert.Single(
            PdfLogicalTableValueAnalysis.Analyze(
                new[] { "Termin" },
                rows,
                new PdfLogicalTableValueAnalysisOptions {
                    NumericCulture = culture,
                    DateTimeCulture = culture
                }));

        Assert.Equal(PdfLogicalTableValueKind.Text, numericOnly.Kind);
        Assert.Equal(1D, numericOnly.Confidence);
        Assert.Equal(PdfLogicalTableValueKind.DateTime, localizedDates.Kind);
    }

    [Fact]
    public void Analyze_DoesNotTreatBareYearsAsLocalizedDates() {
        IReadOnlyList<IReadOnlyList<string>> rows = new[] {
            (IReadOnlyList<string>) new[] { "2025" },
            new[] { "2026" }
        };

        PdfLogicalTableValueProfile profile = Assert.Single(
            PdfLogicalTableValueAnalysis.Analyze(
                new[] { "Rok" },
                rows,
                new PdfLogicalTableValueAnalysisOptions {
                    DateTimeCulture = System.Globalization.CultureInfo.GetCultureInfo("pl-PL")
                }));

        Assert.Equal(PdfLogicalTableValueKind.Number, profile.Kind);
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
        Assert.True(PdfLogicalTableValueParser.TryParsePercentage("٢٥٪", null, out decimal arabic));
        Assert.Equal(0.25M, arabic);
        Assert.True(PdfLogicalTableValueParser.TryParsePercentage("１００％", null, out decimal fullWidth));
        Assert.Equal(1M, fullWidth);
    }

    [Fact]
    public void Analyze_PreservesConsistentCurrencyAffixesAsTypedEvidence() {
        IReadOnlyList<IReadOnlyList<string>> rows = new[] {
            (IReadOnlyList<string>) new[] { "25.00 PLN", "$1,234.50" },
            new[] { "75.25 PLN", "$2,345.75" }
        };

        IReadOnlyList<PdfLogicalTableValueProfile> profiles = PdfLogicalTableValueAnalysis.Analyze(
            new[] { "ISO", "Symbol" },
            rows,
            new PdfLogicalTableValueAnalysisOptions {
                NumericCulture = System.Globalization.CultureInfo.InvariantCulture
            });

        Assert.All(profiles, static profile => Assert.Equal(PdfLogicalTableValueKind.Currency, profile.Kind));
        Assert.Equal(new[] { "PLN", "$" }, profiles.Select(static profile => profile.CurrencyToken));
        Assert.Equal(
            new PdfLogicalCurrencyAffixPosition?[] {
                PdfLogicalCurrencyAffixPosition.Suffix,
                PdfLogicalCurrencyAffixPosition.Prefix
            },
            profiles.Select(static profile => profile.CurrencyAffixPosition));
        Assert.Equal(new bool?[] { true, false }, profiles.Select(static profile => profile.CurrencyAffixUsesSpacing));
        Assert.True(PdfLogicalTableValueParser.TryParseCurrency(
            "25.00 PLN",
            System.Globalization.CultureInfo.InvariantCulture,
            out decimal isoValue,
            out string isoToken));
        Assert.Equal(25M, isoValue);
        Assert.Equal("PLN", isoToken);
        Assert.True(PdfLogicalTableValueParser.TryParseCurrency(
            "10,50 zł",
            System.Globalization.CultureInfo.GetCultureInfo("pl-PL"),
            out decimal localizedValue,
            out string localizedToken));
        Assert.Equal(10.50M, localizedValue);
        Assert.Equal("zł", localizedToken);
    }

    [Theory]
    [InlineData("XCG 100.00", PdfLogicalCurrencyAffixPosition.Prefix)]
    [InlineData("100.00 XCG", PdfLogicalCurrencyAffixPosition.Suffix)]
    public void TryParseCurrency_RecognizesTheCurrentCaribbeanGuilderCode(
        string value,
        PdfLogicalCurrencyAffixPosition expectedPosition) {
        Assert.True(PdfLogicalTableValueParser.TryParseCurrency(
            value,
            System.Globalization.CultureInfo.InvariantCulture,
            out decimal parsed,
            out string token,
            out PdfLogicalCurrencyAffixPosition position,
            out bool usesSpacing));

        Assert.Equal(100M, parsed);
        Assert.Equal("XCG", token);
        Assert.Equal(expectedPosition, position);
        Assert.True(usesSpacing);
    }

    [Fact]
    public void Analyze_KeepsMixedCurrencyAffixPlacementAsText() {
        IReadOnlyList<IReadOnlyList<string>> rows = new[] {
            (IReadOnlyList<string>) new[] { "$25.00" },
            new[] { "75.00$" }
        };

        PdfLogicalTableValueProfile profile = Assert.Single(
            PdfLogicalTableValueAnalysis.Analyze(new[] { "Amount" }, rows));

        Assert.Equal(PdfLogicalTableValueKind.Text, profile.Kind);
        Assert.Null(profile.CurrencyToken);
        Assert.Null(profile.CurrencyAffixPosition);
        Assert.Null(profile.CurrencyAffixUsesSpacing);
    }

    [Theory]
    [InlineData("𞋿12.50", PdfLogicalCurrencyAffixPosition.Prefix)]
    [InlineData("12.50𞋿", PdfLogicalCurrencyAffixPosition.Suffix)]
    public void TryParseCurrency_PreservesSupplementaryUnicodeCurrencySymbols(
        string value,
        PdfLogicalCurrencyAffixPosition expectedPosition) {
        Assert.True(PdfLogicalTableValueParser.TryParseCurrency(
            value,
            System.Globalization.CultureInfo.InvariantCulture,
            out decimal parsed,
            out string token,
            out PdfLogicalCurrencyAffixPosition position,
            out bool usesSpacing));

        Assert.Equal(12.5M, parsed);
        Assert.Equal("𞋿", token);
        Assert.Equal(expectedPosition, position);
        Assert.False(usesSpacing);
    }

    [Theory]
    [InlineData("$12", 0)]
    [InlineData("$12.5", 1)]
    [InlineData("$12.500", 3)]
    [InlineData("12.125 KWD", 3)]
    public void TryParseCurrency_ReportsVisibleFractionalPrecision(string value, int expectedDecimalPlaces) {
        Assert.True(PdfLogicalTableValueParser.TryParseCurrency(
            value,
            System.Globalization.CultureInfo.InvariantCulture,
            out _,
            out _,
            out _,
            out _,
            out int decimalPlaces));

        Assert.Equal(expectedDecimalPlaces, decimalPlaces);
    }

    [Theory]
    [InlineData("-$1,234.00", -1234D, PdfLogicalCurrencyAffixPosition.Prefix, false)]
    [InlineData("($1,234.00)", -1234D, PdfLogicalCurrencyAffixPosition.Prefix, false)]
    [InlineData("-1,234.00 USD", -1234D, PdfLogicalCurrencyAffixPosition.Suffix, true)]
    [InlineData("(1,234.00 USD)", -1234D, PdfLogicalCurrencyAffixPosition.Suffix, true)]
    [InlineData("$-1,234.00", -1234D, PdfLogicalCurrencyAffixPosition.Prefix, false)]
    [InlineData("$1,234.00-", -1234D, PdfLogicalCurrencyAffixPosition.Prefix, false)]
    [InlineData("1,234.00-$", -1234D, PdfLogicalCurrencyAffixPosition.Suffix, false)]
    [InlineData("1,234.00 USD-", -1234D, PdfLogicalCurrencyAffixPosition.Suffix, true)]
    public void TryParseCurrency_RecognizesSignedAndAccountingAffixes(
        string value,
        double expected,
        PdfLogicalCurrencyAffixPosition expectedPosition,
        bool expectedSpacing) {
        Assert.True(PdfLogicalTableValueParser.TryParseCurrency(
            value,
            System.Globalization.CultureInfo.InvariantCulture,
            out decimal parsed,
            out _,
            out PdfLogicalCurrencyAffixPosition position,
            out bool usesSpacing));

        Assert.Equal((decimal)expected, parsed);
        Assert.Equal(expectedPosition, position);
        Assert.Equal(expectedSpacing, usesSpacing);
    }

    [Fact]
    public void Analyze_KeepsSignedCurrencyColumnsTyped() {
        IReadOnlyList<IReadOnlyList<string>> rows = new[] {
            (IReadOnlyList<string>) new[] { "$1,234.00" },
            new[] { "-$20.50" },
            new[] { "($5.25)" },
            new[] { "$6.75-" }
        };

        PdfLogicalTableValueProfile profile = Assert.Single(PdfLogicalTableValueAnalysis.Analyze(
            new[] { "Amount" },
            rows,
            new PdfLogicalTableValueAnalysisOptions {
                NumericCulture = System.Globalization.CultureInfo.InvariantCulture
            }));

        Assert.Equal(PdfLogicalTableValueKind.Currency, profile.Kind);
        Assert.Equal("$", profile.CurrencyToken);
        Assert.Equal(PdfLogicalCurrencyAffixPosition.Prefix, profile.CurrencyAffixPosition);
    }

    [Fact]
    public void TryParseCurrency_RecognizesCurrentBmpCurrencySymbolsAcrossTargets() {
        Assert.True(PdfLogicalTableValueParser.TryParseCurrency(
            "\u20C112.50",
            System.Globalization.CultureInfo.InvariantCulture,
            out decimal parsed,
            out string token));

        Assert.Equal(12.5M, parsed);
        Assert.Equal("\u20C1", token);
    }

    [Theory]
    [InlineData("123kg")]
    [InlineData("123 ABC")]
    [InlineData("123 ABCD")]
    [InlineData("usd 123")]
    [InlineData("ABC text")]
    public void TryParseCurrency_RejectsUnitsAndUnstructuredText(string value) {
        Assert.False(PdfLogicalTableValueParser.TryParseCurrency(value, null, out _, out _));
    }
}
