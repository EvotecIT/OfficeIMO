using OfficeIMO.Excel.Legacy;
using OfficeIMO.Excel.Html;
using OfficeIMO.Excel.Csv;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Excel.OpenDocument;
using OfficeIMO.Reader;
using OfficeIMO.Reader.Excel;

namespace OfficeIMO.LegacyImport.Tests;

public sealed class LegacySpreadsheetImportTests {
    public static IEnumerable<object[]> Families() {
        yield return new object[] { LegacyFixtureFactory.Wk(), "archive.wk1", LegacySpreadsheetFormat.Lotus123 };
        yield return new object[] { LegacyFixtureFactory.Wk(0x20, 0x51), "archive.wq1", LegacySpreadsheetFormat.QuattroPro };
        yield return new object[] { LegacyFixtureFactory.CompoundSheet(), "archive.qpw", LegacySpreadsheetFormat.QuattroPro };
        yield return new object[] { LegacyFixtureFactory.Multiplan(), "archive.mp", LegacySpreadsheetFormat.Multiplan };
        yield return new object[] { LegacyFixtureFactory.Wk(0x04, 0x04, includeFormulaAndChart: false), "archive.wks", LegacySpreadsheetFormat.MicrosoftWorks };
        yield return new object[] { LegacyFixtureFactory.CompoundSheet(), "archive.xlr", LegacySpreadsheetFormat.MicrosoftWorks };
    }

    [Theory]
    [MemberData(nameof(Families))]
    public void DetectsAndImportsEveryBoundedFamily(byte[] source, string sourceName, LegacySpreadsheetFormat expected) {
        using LegacySpreadsheetImportResult result = LegacySpreadsheetImporter.Import(source, new LegacySpreadsheetImportOptions { SourceName = sourceName });
        Assert.Equal(expected, result.Detection.Format);
        Assert.True(result.Report.RecoveredItemCount > 0);
        Assert.NotEmpty(result.Document.Sheets);
    }

    [Fact]
    public void WkRecordsRecoverCachedValuesAlignmentAndChartMetadata() {
        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(LegacyFixtureFactory.Wk(), new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" });
        Assert.Equal(OfficeLegacyImportQuality.Structured, imported.Report.Quality);
        Assert.Single(imported.Charts);
        Assert.DoesNotContain(imported.Report.Findings, finding => finding.Code == "WK_FORMULA_CACHED_FALLBACK");
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "LEGACY_CHART_METADATA_ONLY");
        Assert.Equal("($A$1+$B$1)", Assert.Single(imported.Cells, cell => cell.Row == 1 && cell.Column == 3).Formula);
        Assert.Equal(84d, Assert.Single(imported.Cells, cell => cell.Row == 1 && cell.Column == 3).CachedValue);
        Assert.Equal("Input", Assert.Single(imported.Names).Name);
        Assert.Equal("Input", Assert.Single(imported.Names).ProjectedName);
        Assert.Contains(imported.Document.CreateInspectionSnapshot().NamedRanges, name => name.Name == "Input");
        var sheet = Assert.Single(imported.Document.CreateInspectionSnapshot().Worksheets);
        Assert.Equal("Name", Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1).Value);
        Assert.Equal("left", Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 1).Style?.HorizontalAlignment);
        Assert.Equal(OfficeIMO.Excel.ExcelHorizontalAlignment.Left, Assert.Single(imported.Cells, cell => cell.Row == 1 && cell.Column == 1).Alignment);
        Assert.Equal(42, Convert.ToInt32(Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 2).Value));
        var formulaCell = Assert.Single(sheet.Cells, cell => cell.Row == 1 && cell.Column == 3);
        Assert.Equal("($A$1+$B$1)", formulaCell.Formula);
        Assert.Equal(84d, Convert.ToDouble(formulaCell.Value));
        Assert.Equal("84", imported.Document.Sheets[0].CellAt(1, 3).GetValue().CachedText);
    }

    [Fact]
    public void WkLabelsPreserveWhitespaceAndDoNotInventActiveContentFromVisibleText() {
        const string label = "  https://example.invalid macro  ";
        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.Wk(includeFormulaAndChart: false, label: label),
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" });

        Assert.Equal(label, Assert.Single(imported.Cells, cell => cell.Row == 1 && cell.Column == 1).CachedValue);
        Assert.Equal(OfficeLegacyInertContentKind.None, imported.Report.InertContent);
        Assert.DoesNotContain(imported.Report.Findings, finding =>
            finding.Code == "LEGACY_EXTERNAL_LINK_INERT" || finding.Code == "LEGACY_MACRO_INERT");
    }

    [Fact]
    public void WkLabelsRequireTheirRecordTerminator() {
        Assert.Throws<InvalidDataException>(() => LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.Wk(includeFormulaAndChart: false, terminateLabel: false),
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" }));
    }

    [Fact]
    public void StructuredWkLabelsRespectTheExcelCellLimitWithLossDiagnostics() {
        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.Wk(includeFormulaAndChart: false, label: new string('L', 40_000)),
            new LegacySpreadsheetImportOptions {
                SourceName = "archive.wk1",
                Limits = new OfficeLegacyImportLimits { MaxTextCharacters = 50_000 }
            });

        string value = Assert.IsType<string>(Assert.Single(imported.Cells, cell => cell.Row == 1 && cell.Column == 1).CachedValue);
        Assert.Equal(32_767, value.Length);
        Assert.Equal("1", imported.Metadata["TruncatedStructuredCellCount"]);
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "LEGACY_SHEET_CELL_TEXT_TRUNCATED");
    }

    [Fact]
    public void WkFormulaFallbackAndRecordBoundsAreExplicit() {
        using LegacySpreadsheetImportResult fallback = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.Wk(formulaTokens: new byte[] { 0xFE, 0x03 }),
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" });
        LegacySpreadsheetCellContent formula = Assert.Single(fallback.Cells, cell => cell.Row == 1 && cell.Column == 3);
        Assert.Null(formula.Formula);
        Assert.Equal(84d, formula.CachedValue);
        Assert.Contains(fallback.Report.Findings, finding => finding.Code == "WK_FORMULA_CACHED_FALLBACK");

        Assert.Throws<InvalidDataException>(() => LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.Wk(declaredFormulaLength: 99),
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" }));

        var expansiveFormula = new List<byte>();
        for (short value = 0; value < 6; value++) { expansiveFormula.Add(0x05); expansiveFormula.Add((byte)value); expansiveFormula.Add(0); }
        expansiveFormula.AddRange(Enumerable.Repeat((byte)0x09, 5));
        expansiveFormula.Add(0x03);
        using LegacySpreadsheetImportResult bounded = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.Wk(formulaTokens: expansiveFormula.ToArray()),
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk1", Limits = new OfficeLegacyImportLimits { MaxItems = 10 } });
        Assert.Null(Assert.Single(bounded.Cells, cell => cell.Row == 1 && cell.Column == 3).Formula);
        Assert.Contains(bounded.Metadata.Values, value => value.Contains("expression-node limit", StringComparison.Ordinal));
    }

    [Fact]
    public void LotusErrFormulaRetainsItsCachedValueAsUnsupported() {
        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.Wk(formulaTokens: new byte[] { 0x20, 0x03 }),
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" });

        LegacySpreadsheetCellContent formula = Assert.Single(imported.Cells, cell => cell.Row == 1 && cell.Column == 3);
        Assert.Null(formula.Formula);
        Assert.Equal(84d, formula.CachedValue);
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "WK_FORMULA_CACHED_FALLBACK");

        using LegacySpreadsheetImportResult notAvailable = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.Wk(formulaTokens: new byte[] { 0x1F, 0x03 }),
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" });
        Assert.Equal("NA()", Assert.Single(notAvailable.Cells, cell => cell.Row == 1 && cell.Column == 3).Formula);
    }

    [Fact]
    public void WkFormulaTextSharesTheWorkbookCharacterBudget() {
        const int sourceNameAndLabelCharacters = 9;
        const int oneFormulaCharacters = 13;
        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.WkWithRepeatedFormulas(),
            new LegacySpreadsheetImportOptions {
                SourceName = "archive.wk1",
                Limits = new OfficeLegacyImportLimits { MaxTextCharacters = sourceNameAndLabelCharacters + oneFormulaCharacters }
            });

        Assert.Single(imported.Cells, cell => cell.Formula != null);
        Assert.Single(imported.Cells, cell => cell.Formula == null && cell.CachedValue is double);
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "WK_FORMULA_CACHED_FALLBACK");
        Assert.True(imported.Names.Sum(name => name.Name.Length) +
            imported.Cells.Sum(cell => (cell.CachedValue as string)?.Length ?? 0) +
            imported.Cells.Sum(cell => cell.Formula?.Length ?? 0) <= sourceNameAndLabelCharacters + oneFormulaCharacters);
    }

    [Fact]
    public void WkFallbackMetadataUsesBoundedSamplesAndAggregateCounts() {
        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.WkWithRepeatedFallbackMetadata(20),
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" });

        Assert.Equal("20", imported.Metadata["FormulaFallbackCount"]);
        Assert.Equal("20", imported.Metadata["UnresolvedNameCount"]);
        Assert.Equal(16, imported.Metadata.Keys.Count(key => key.StartsWith("FormulaFallback.Sample.", StringComparison.Ordinal)));
        Assert.Equal(16, imported.Metadata.Keys.Count(key => key.StartsWith("UnresolvedName.Sample.", StringComparison.Ordinal)));
    }

    [Fact]
    public void WkBlankCellsAndUnknownRecordsCannotDisappearFromNoLossClaims() {
        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.Wk(includeBlank: true, extraRecordType: 0x7777),
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" });
        Assert.Contains(imported.Cells, cell => cell.Row == 1 && cell.Column == 4 && cell.CachedValue == null);
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "WK_RECORDS_UNSUPPORTED");
        Assert.Throws<InvalidOperationException>(() => imported.Report.RequireStructuredNoLoss());
    }

    [Fact]
    public void NamedRangeProjectionIsStrictCollisionSafeAndObservable() {
        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.WkWithNames("Input", "input", "A1"),
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" });
        Assert.Equal("Input", imported.Names[0].ProjectedName);
        Assert.Null(imported.Names[1].ProjectedName);
        Assert.Null(imported.Names[2].ProjectedName);
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "WK_NAME_COLLISION");
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "WK_NAME_INVALID");
        Assert.Single(imported.Document.CreateInspectionSnapshot().NamedRanges);
    }

    [Fact]
    public void Wk1NamesUseSixteenBitColumnsOnTheSingleSourceSheet() {
        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.WkWithWideColumnName(),
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk1", RequireStructured = true });
        LegacySpreadsheetNameContent projected = Assert.Single(imported.Names);
        Assert.Equal("Sheet1", projected.SheetName);
        Assert.Equal(257, projected.FirstColumn);
        Assert.Equal(258, projected.LastColumn);
        Assert.Contains(imported.Document.ListNamedRanges(), name => name.Name == "Wide" &&
            name.Reference.Contains("$IW$1:$IX$1", StringComparison.Ordinal));
    }

    [Fact]
    public void QuattroFormulaEnvelopeIsValidatedEvenWhenTokensRemainInert() {
        Assert.Throws<InvalidDataException>(() => LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.Wq2WithTruncatedFormulaEnvelope(),
            new LegacySpreadsheetImportOptions { SourceName = "archive.wq2" }));
    }

    [Fact]
    public void WkSheetIdentifiersProjectToSeparateWorksheets() {
        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(LegacyFixtureFactory.WkMultiSheet(), new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" });
        Assert.Equal(2, imported.Document.Sheets.Count);
        Assert.Contains(imported.Cells, cell => cell.SheetName == "Sheet1" && Convert.ToInt32(cell.CachedValue) == 1);
        Assert.Contains(imported.Cells, cell => cell.SheetName == "Sheet2" && Convert.ToInt32(cell.CachedValue) == 2);
    }

    [Fact]
    public void WkParserRejectsRecordsAfterEofAndBoundsMetadataText() {
        byte[] trailingCell = { 0x0D, 0x00, 0x07, 0x00, 0x00, 0x03, 0x00, 0x00, 0x00, 0x63, 0x00 };
        byte[] source = LegacyFixtureFactory.Wk().Concat(trailingCell).ToArray();
        Assert.Throws<InvalidDataException>(() => LegacySpreadsheetImporter.Import(
            source,
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" }));

        byte[] eofPayload = LegacyFixtureFactory.Wk(includeFormulaAndChart: false);
        eofPayload[eofPayload.Length - 2] = 1;
        Assert.Throws<InvalidDataException>(() => LegacySpreadsheetImporter.Import(
            eofPayload.Concat(new byte[] { 0 }).ToArray(),
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" }));

        using LegacySpreadsheetImportResult padded = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.Wk(includeFormulaAndChart: false).Concat(new byte[] { 0, 0, 0x1A }).ToArray(),
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk1", RequireStructured = true });
        Assert.Equal(OfficeLegacyImportQuality.Structured, padded.Report.Quality);

        Assert.Throws<InvalidDataException>(() => LegacySpreadsheetImporter.Import(LegacyFixtureFactory.Wk(), new LegacySpreadsheetImportOptions {
            SourceName = "archive.wk1",
            Limits = new OfficeLegacyImportLimits { MaxTextCharacters = 3 }
        }));

        using LegacySpreadsheetImportResult formatted = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.Wk(cellFormat: 0x60),
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" });
        Assert.Contains(formatted.Report.Findings, finding => finding.Code == "WK_CELL_FORMAT_PARTIAL");
    }

    [Fact]
    public void WkStructuredImportRequiresTheValidatedFamilyBofPayload() {
        Assert.Throws<InvalidDataException>(() => LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.Wk(product0: 0x99, product1: 0x99),
            new LegacySpreadsheetImportOptions {
                FormatHint = LegacySpreadsheetFormat.Lotus123,
                RequireStructured = true
            }));

        using LegacySpreadsheetImportResult salvage = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.Wk(product0: 0x99, product1: 0x99),
            new LegacySpreadsheetImportOptions { FormatHint = LegacySpreadsheetFormat.Lotus123 });
        Assert.Equal(OfficeLegacyImportQuality.Salvage, salvage.Report.Quality);
        Assert.Equal("lotus-1-2-3-later-salvage", salvage.Report.SourceFormatId);

        Assert.Throws<InvalidDataException>(() => LegacySpreadsheetImporter.Detect(
            LegacyFixtureFactory.Wk(product0: 0x99, product1: 0x99),
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" }));
    }

    [Fact]
    public void LaterLotusEnvelopeRequiresCorroboratingFamilyEvidence() {
        byte[] laterEnvelope = new byte[] { 0x00, 0x00, 0x1A, 0x00 }
            .Concat(Encoding.ASCII.GetBytes("Recoverable\tLotus text\n"))
            .ToArray();

        Assert.Throws<InvalidDataException>(() => LegacySpreadsheetImporter.Detect(laterEnvelope));
        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(
            laterEnvelope,
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk3" });
        Assert.Equal(LegacySpreadsheetFormat.Lotus123, imported.Detection.Format);
        Assert.Equal(OfficeLegacyImportQuality.Salvage, imported.Report.Quality);
    }

    [Fact]
    public void SalvagedTextCellsRespectTheExcelCellLimitWithLossDiagnostics() {
        byte[] source = new byte[] { 0x08, 0xE7 }
            .Concat(Encoding.ASCII.GetBytes(new string('A', 40_000) + "\n"))
            .ToArray();

        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(
            source,
            new LegacySpreadsheetImportOptions {
                FormatHint = LegacySpreadsheetFormat.Multiplan,
                Limits = new OfficeLegacyImportLimits { MaxTextCharacters = 50_000 }
            });

        string value = Assert.IsType<string>(Assert.Single(imported.Cells).CachedValue);
        Assert.Equal(32_767, value.Length);
        Assert.Equal("1", imported.Metadata["TruncatedSalvageCellCount"]);
        Assert.Contains(imported.Report.Findings, finding => finding.Code == "LEGACY_SHEET_CELL_TEXT_TRUNCATED");
        Assert.Equal(value, imported.Document.Sheets[0].CellAt(1, 1).GetValue().CachedText);
    }

    [Fact]
    public void WeakExtensionsAndUninspectableCompoundSecurityDoNotPassSilently() {
        Assert.Throws<InvalidDataException>(() => LegacySpreadsheetImporter.Import(
            Encoding.ASCII.GetBytes("renamed,plain,text"),
            new LegacySpreadsheetImportOptions { SourceName = "renamed.wk1" }));

        using LegacySpreadsheetImportResult compound = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.TruncatedCompoundHeader(),
            new LegacySpreadsheetImportOptions { FormatHint = LegacySpreadsheetFormat.MicrosoftWorks });
        Assert.Contains(compound.Report.Findings, finding => finding.Code == "LEGACY_COMPOUND_INVENTORY_INCOMPLETE");
        Assert.True(compound.Report.HasLoss);

        using LegacySpreadsheetImportResult validCompound = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.CompoundSheet(),
            new LegacySpreadsheetImportOptions { SourceName = "archive.xlr" });
        Assert.DoesNotContain(validCompound.Report.Findings, finding => finding.Code == "LEGACY_COMPOUND_INVENTORY_INCOMPLETE");
    }

    [Fact]
    public void ImportedWorkbookUsesEverySupportedModernOutputOwner() {
        using LegacySpreadsheetImportResult imported = LegacySpreadsheetImporter.Import(LegacyFixtureFactory.Wk(), new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" });
        using var xlsx = new MemoryStream();
        imported.Document.Save(xlsx);
        Assert.True(xlsx.Length > 100);
        Assert.Contains("Name", imported.Document.ToHtml());
        Assert.Contains("Name", imported.Document.Sheets[0].ToCsv());
        Assert.StartsWith("%PDF", Encoding.ASCII.GetString(imported.Document.ToPdf(), 0, 4));
        using var ods = new MemoryStream();
        imported.Document.ToOpenDocument().Save(ods);
        Assert.True(ods.Length > 100);
    }

    [Fact]
    public void ReaderHandlerProjectsLegacyWarningsAndContent() {
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddLegacySpreadsheetHandler().Build();
        using var stream = new MemoryStream(LegacyFixtureFactory.Wk());
        OfficeDocumentReadResult result = reader.ReadDocument(stream, "archive.wk1");
        Assert.Contains(result.Chunks, chunk => chunk.Text.Contains("Name", StringComparison.Ordinal));
        Assert.Contains(OfficeDocumentReaderBuilderExcelExtensions.LegacyHandlerId, result.CapabilitiesUsed);
        Assert.Contains(result.Chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>()), warning => warning.Contains("Legacy import quality", StringComparison.Ordinal));
    }

    [Fact]
    public void ImportHonorsCancellation() {
        Assert.Throws<OperationCanceledException>(() => LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.Wk(),
            new LegacySpreadsheetImportOptions { SourceName = "archive.wk1" },
            new CancellationToken(canceled: true)));
    }

    [Fact]
    public void ReaderRegistrationCapturesLegacyImportOptions() {
        var options = new LegacySpreadsheetImportOptions {
            FormatHint = LegacySpreadsheetFormat.Lotus123,
            Limits = new OfficeLegacyImportLimits { MaxInputBytes = 1024 }
        };
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddLegacySpreadsheetHandler(options).Build();
        options.FormatHint = LegacySpreadsheetFormat.Multiplan;
        options.Limits.MaxInputBytes = 1;

        using var stream = new MemoryStream(LegacyFixtureFactory.Wk());
        OfficeDocumentReadResult result = reader.ReadDocument(stream, "archive.wk1");
        Assert.Contains(result.Chunks.SelectMany(chunk => chunk.Warnings ?? Array.Empty<string>()),
            warning => warning.Contains("lotus-1-2-3", StringComparison.Ordinal));
    }

    [Fact]
    public void ShortSalvagePrefixesRequireCorroboratingSourceEvidence() {
        Assert.Throws<InvalidDataException>(() => LegacySpreadsheetImporter.Detect(LegacyFixtureFactory.Multiplan()));
        byte[] works = new byte[] { 0xFF, 0x00, 0x02 }.Concat(Encoding.ASCII.GetBytes("Recoverable text")).ToArray();
        Assert.Throws<InvalidDataException>(() => LegacySpreadsheetImporter.Detect(works));

        Assert.Equal(LegacySpreadsheetFormat.Multiplan,
            LegacySpreadsheetImporter.Detect(LegacyFixtureFactory.Multiplan(), new LegacySpreadsheetImportOptions { SourceName = "archive.mp" }).Format);
        Assert.Equal(LegacySpreadsheetFormat.MicrosoftWorks,
            LegacySpreadsheetImporter.Detect(works, new LegacySpreadsheetImportOptions { SourceName = "archive.wks" }).Format);
    }

    [Fact]
    public void ExactStructuredBofSignaturesDoNotRequireSourceNames() {
        using LegacySpreadsheetImportResult lotus = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.Wk(includeFormulaAndChart: false),
            new LegacySpreadsheetImportOptions { RequireStructured = true });
        using LegacySpreadsheetImportResult works = LegacySpreadsheetImporter.Import(
            LegacyFixtureFactory.Wk(0x04, 0x04, includeFormulaAndChart: false),
            new LegacySpreadsheetImportOptions { RequireStructured = true });

        Assert.Equal(LegacySpreadsheetFormat.Lotus123, lotus.Detection.Format);
        Assert.Equal(LegacySpreadsheetFormat.MicrosoftWorks, works.Detection.Format);
    }

    [Fact]
    public void ReaderLimitCannotRaiseConfiguredLegacySpreadsheetLimit() {
        byte[] source = LegacyFixtureFactory.Wk();
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddLegacySpreadsheetHandler(new LegacySpreadsheetImportOptions {
                Limits = new OfficeLegacyImportLimits { MaxInputBytes = source.Length - 1 }
            })
            .Build();

        using var stream = new MemoryStream(source);
        Assert.Throws<InvalidDataException>(() => reader.ReadDocument(stream, "archive.wk1",
            new ReaderOptions { MaxInputBytes = source.Length + 100L }));
    }

    [Fact]
    public void LegacySpreadsheetHandlerUsesConfiguredLimitBeforeBufferingNonSeekableStreams() {
        byte[] source = LegacyFixtureFactory.Wk();
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
            .AddLegacySpreadsheetHandler(new LegacySpreadsheetImportOptions {
                Limits = new OfficeLegacyImportLimits { MaxInputBytes = source.Length - 1 }
            })
            .Build();
        using var stream = new NonSeekableStream(source);

        Assert.Throws<IOException>(() => reader.ReadDocument(stream, "archive.wk1"));
    }

    private sealed class NonSeekableStream : Stream {
        private readonly MemoryStream _inner;
        internal NonSeekableStream(byte[] data) => _inner = new MemoryStream(data, writable: false);
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position { get => throw new NotSupportedException(); set => throw new NotSupportedException(); }
        public override void Flush() { }
        public override int Read(byte[] buffer, int offset, int count) => _inner.Read(buffer, offset, count);
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        protected override void Dispose(bool disposing) {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
