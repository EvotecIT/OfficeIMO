using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {


    [Theory]
    [InlineData("xl%2Fworkbook.bin")]
    [InlineData("xl%5cworkbook.bin")]
    public void ExcelXlsbRejectsPercentEncodedPackageSeparators(string target) {
        byte[] package = CreateWave33XlsbProvenancePackage(signed: false, officeDocumentTarget: target);

        Assert.ThrowsAny<Exception>(() => ExcelDocument.RemoveProvenance(package, "workbook.xlsb"));
    }
}
