using System.Reflection;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.LegacyXls.Projection;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void ValuesModeWorkbookMergeDoesNotResolveOrCopySourceDefinedNames() {
        using var sourceStream = new MemoryStream();
        using (ExcelDocument source = ExcelDocument.Create(sourceStream)) {
            ExcelSheet data = source.AddWorksheet("Data");
            data.CellValue(1, 1, 21);
            data.CellFormula(1, 2, "TaxRate*2");
            source.SetNamedRange("TaxRate", "A1", data, save: false);
            source.Save(sourceStream);
        }

        sourceStream.Position = 0;
        using ExcelDocument sourceDocument = ExcelDocument.Load(sourceStream, new ExcelLoadOptions {
            AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly
        });
        using ExcelDocument target = ExcelDocument.Create(new MemoryStream());
        target.AddWorksheet("Existing");

        target.MergeWorkbookFrom(sourceDocument, new ExcelWorkbookMergeOptions {
            CopyMode = ExcelWorksheetCopyMode.Values
        });

        Assert.Null(target.WorkbookPartRoot.Workbook.DefinedNames);
        Assert.Empty(target["Data"].WorksheetPart.Worksheet.Descendants<CellFormula>());
    }

    [Fact]
    public void LegacyExternalReferenceFilterTracksQuotedQualifiersIndependentlyFromStringLiterals() {
        byte[] workbookStream = LegacyXlsTestWorkbookBuilder.CreateFormulaExternalWorkbookReferenceWorkbookStream();
        byte[] compound = LegacyXlsCompoundTestBuilder.CreateWorkbookCompoundFile(workbookStream);
        LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(compound, new LegacyXlsImportOptions());
        MethodInfo method = typeof(LegacyXlsWorkbookProjector).GetMethod(
            "ReferencesExternalWorkbook",
            BindingFlags.Static | BindingFlags.NonPublic)!;

        bool external = (bool)method.Invoke(null, new object[] {
            workbook,
            "'Decoy\"Sheet'!A1+'[Budget.xls]Jan'!A1"
        })!;
        bool literalOnly = (bool)method.Invoke(null, new object[] {
            workbook,
            "\"[Budget.xls]Jan\"&A1"
        })!;

        Assert.True(external);
        Assert.False(literalOnly);
    }

    [Fact]
    public void LegacyExternalReferenceFilterHandlesClosingBracketInWorkbookName() {
        Type matcherType = typeof(LegacyXlsWorkbookProjector).GetNestedType(
            "ExternalWorkbookReferenceMatcher",
            BindingFlags.NonPublic)!;
        var reference = new LegacyXlsExternalReference(
            LegacyXlsExternalReferenceKind.ExternalWorkbook,
            "Budget]2025.xls",
            new[] { "Jan" },
            1);
        object matcher = Activator.CreateInstance(
            matcherType,
            BindingFlags.Instance | BindingFlags.NonPublic,
            binder: null,
            args: new object[] { new[] { reference } },
            culture: null)!;
        MethodInfo matches = matcherType.GetMethod(
            "ReferencesExternalWorkbook",
            BindingFlags.Instance | BindingFlags.NonPublic)!;

        Assert.True((bool)matches.Invoke(matcher, new object[] { "'[Budget]2025.xls]Jan'!A1" })!);
    }

    [Fact]
    public void TemporaryPackageStagingStopsBeforeExceedingItsConfiguredLimit() {
        using var inner = new MemoryStream();
        using var staging = new ExcelBoundedSeekableStream(inner, maximumBytes: 8, leaveOpen: true);
        staging.Write(new byte[8], 0, 8);

        IOException exception = Assert.Throws<IOException>(() => staging.WriteByte(1));

        Assert.Equal(8, inner.Length);
        Assert.Contains("8-byte temporary package limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void ConversionSaveOptionsPreserveTheTemporaryPackageLimit() {
        var options = new ExcelSaveOptions { MaxTemporaryPackageBytes = 1_024 };

        ExcelSaveOptions copy = options.WithLossPolicy(ExcelConversionLossPolicy.Allow);

        Assert.Equal(1_024, copy.MaxTemporaryPackageBytes);
    }
}
