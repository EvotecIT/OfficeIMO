using System.Reflection;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfTableConversionApiContracts {
    [Theory]
    [InlineData(typeof(PdfExcelTableConverterExtensions), "SaveAsExcelFromPdfTables", "ToExcelBytesFromPdfTables", "SavePdfTablesAsExcel", "ToExcelTableWorkbookBytes")]
    [InlineData(typeof(PdfWordTableConverterExtensions), "SaveAsWordFromPdfTables", "ToWordBytesFromPdfTables", "SavePdfTablesAsWord", "ToWordTableDocumentBytes")]
    [InlineData(typeof(PowerPointPdfConverterExtensions), "SaveAsPowerPointFromPdfTables", "ToPowerPointBytesFromPdfTables", "SavePdfTablesAsPowerPoint", "ToPowerPointTablePresentationBytes")]
    public void TableConversionApisUseDestinationFirstNames(
        Type converterType,
        string saveName,
        string bytesName,
        string removedSaveName,
        string removedBytesName) {
        string[] methodNames = converterType
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Select(method => method.Name)
            .ToArray();

        Assert.Equal(5, methodNames.Count(name => name == saveName));
        Assert.Single(methodNames, name => name == bytesName);
        Assert.DoesNotContain(removedSaveName, methodNames);
        Assert.DoesNotContain(removedBytesName, methodNames);
    }
}
