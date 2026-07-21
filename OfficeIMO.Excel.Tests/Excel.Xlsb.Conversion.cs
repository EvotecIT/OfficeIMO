using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Excel {
    [Fact]
    public void Xlsb_ConversionReportsEveryUnprojectedRecordBeforeCreatingXlsx() {
        string sourcePath = Path.Combine(
            AppContext.BaseDirectory,
            "Documents",
            "XlsbCorpus",
            "excel-generated",
            "basic-values-formula.xlsb");
        string blockedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
        string allowedPath = Path.Combine(_directoryWithFiles, Guid.NewGuid().ToString("N") + ".xlsx");
        int preservedRecordCount;
        using (ExcelDocument source = ExcelDocument.Load(sourcePath)) {
            preservedRecordCount = source.XlsbPreservedRecords.Count;
        }
        Assert.True(preservedRecordCount > 0);

        ExcelDocumentConversionException blocked = Assert.Throws<ExcelDocumentConversionException>(() =>
            ExcelDocument.Convert(sourcePath, blockedPath));

        Assert.Equal(ExcelDocumentConversionFailureReason.DataLossBlocked, blocked.Reason);
        Assert.False(File.Exists(blockedPath));
        OfficeCompatibilityFinding[] recordFindings = blocked.Result.Report.Compatibility.Findings
            .Where(finding => finding.Code.StartsWith("Excel.Xlsb.UnprojectedRecord.", StringComparison.Ordinal))
            .ToArray();
        Assert.Equal(preservedRecordCount, recordFindings.Length);
        Assert.All(recordFindings, finding => {
            Assert.Equal(OfficeCompatibilityState.Blocked, finding.State);
            Assert.True(finding.RepresentsLoss);
            Assert.NotNull(finding.SourceLocation);
            Assert.True((finding.Impact & OfficeCompatibilityImpact.Carrier) != 0);
        });

        ExcelDocumentConversionResult allowed = ExcelDocument.Convert(
            sourcePath,
            allowedPath,
            new ExcelDocumentConversionOptions { LossPolicy = ExcelConversionLossPolicy.Allow });

        Assert.Equal(OfficeCompatibilityMode.BestEffort, allowed.Report.Compatibility.Mode);
        Assert.True(allowed.Report.Compatibility.HasLoss);
        OfficeCompatibilityFinding[] allowedRecordFindings = allowed.Report.Compatibility.Findings
            .Where(finding => finding.Code.StartsWith("Excel.Xlsb.UnprojectedRecord.", StringComparison.Ordinal))
            .ToArray();
        Assert.Equal(preservedRecordCount, allowedRecordFindings.Length);
        Assert.All(allowedRecordFindings, finding =>
            Assert.Equal(OfficeCompatibilityState.Dropped, finding.State));
        Assert.True(File.Exists(allowedPath));
    }
}
