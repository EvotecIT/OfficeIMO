using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ExcelSelectiveWorkbookProtectionTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void RemovingOneRestrictionPreservesTheOtherAndPasswordHashes(bool structure) {
        using var workbook = ExcelDocument.Create();
        workbook.AddSheet("Data");
        workbook.ProtectWorkbook(new ExcelWorkbookProtectionOptions {
            ProtectStructure = true, ProtectWindows = true, Password = "synthetic-protection-password"
        });
        using var baseline = SpreadsheetDocument.Open(new MemoryStream(workbook.ToBytes()), false);
        string hash = baseline.WorkbookPart!.Workbook!.GetFirstChild<WorkbookProtection>()!.WorkbookPassword!.Value!;

        workbook.UnprotectWorkbook(structure: structure, windows: !structure);
        using var saved = SpreadsheetDocument.Open(new MemoryStream(workbook.ToBytes()), false);
        WorkbookProtection protection = Assert.Single(saved.WorkbookPart!.Workbook!.Elements<WorkbookProtection>());
        Assert.Equal(!structure, protection.LockStructure!.Value);
        Assert.Equal(structure, protection.LockWindows!.Value);
        Assert.Equal(hash, protection.WorkbookPassword!.Value);

        workbook.UnprotectWorkbook();
        using var unprotected = SpreadsheetDocument.Open(new MemoryStream(workbook.ToBytes()), false);
        Assert.Empty(unprotected.WorkbookPart!.Workbook!.Elements<WorkbookProtection>());
    }

    [Fact]
    public void SubsetRemovalRetainsRevisionAndModernPasswordMetadata() {
        using var original = ExcelDocument.Create();
        original.AddSheet("Data");
        original.ProtectWorkbook(new ExcelWorkbookProtectionOptions { ProtectStructure = true });
        using var input = new MemoryStream();
        byte[] bytes = original.ToBytes();
        input.Write(bytes, 0, bytes.Length);
        using (var package = SpreadsheetDocument.Open(input, true)) {
            WorkbookProtection protection = package.WorkbookPart!.Workbook!.GetFirstChild<WorkbookProtection>()!;
            protection.LockRevision = true;
            protection.WorkbookAlgorithmName = "SHA-512";
            protection.WorkbookHashValue = "AQIDBA==";
            protection.WorkbookSaltValue = "BQYHCA==";
            protection.WorkbookSpinCount = 100000U;
            protection.RevisionsPassword = "ABCD";
            package.WorkbookPart.Workbook!.Save();
        }
        input.Position = 0;
        using var workbook = ExcelDocument.Load(input);
        workbook.UnprotectWorkbook(structure: true, windows: false);
        using var saved = SpreadsheetDocument.Open(new MemoryStream(workbook.ToBytes()), false);
        WorkbookProtection retained = Assert.Single(saved.WorkbookPart!.Workbook!.Elements<WorkbookProtection>());
        Assert.True(retained.LockRevision!.Value);
        Assert.Equal("SHA-512", retained.WorkbookAlgorithmName!.Value);
        Assert.Equal("AQIDBA==", retained.WorkbookHashValue!.Value);
        Assert.Equal("BQYHCA==", retained.WorkbookSaltValue!.Value);
        Assert.Equal(100000U, retained.WorkbookSpinCount!.Value);
        Assert.Equal("ABCD", retained.RevisionsPassword!.Value);
        Assert.Throws<ArgumentException>(() => workbook.UnprotectWorkbook(false, false));
    }
}
