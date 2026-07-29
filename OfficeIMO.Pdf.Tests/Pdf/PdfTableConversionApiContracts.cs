using System.Reflection;
using System.Threading.Tasks;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfTableConversionApiContracts {
    [Theory]
    [InlineData(typeof(PdfExcelConverterExtensions), "SaveAsExcel", "ToExcelDocument", "ToExcelDocumentResult")]
    [InlineData(typeof(PowerPointPdfConverterExtensions), "SaveAsPowerPoint", "ToPowerPointPresentation", "ToPowerPointPresentationResult")]
    [InlineData(typeof(PdfWordConverterExtensions), "SaveAsWord", "ToWordDocument", "ToWordDocumentResult")]
    [InlineData(typeof(RtfPdfConverterExtensions), "SaveAsRtf", "ToRtfDocument", "ToRtfDocumentResult")]
    [InlineData(typeof(PdfHtmlConverterExtensions), "SaveAsHtml", "ToHtml", "ToHtmlResult")]
    public void ReverseOfficeAdaptersUseTheSameFacadeOnOpenedAndLogicalPdfDocuments(
        Type converterType,
        string saveName,
        string importName,
        string resultName) {
        MethodInfo[] methods = converterType
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .ToArray();
        string[] methodNames = methods.Select(method => method.Name).ToArray();

        Assert.Equal(4, methodNames.Count(name => name == saveName));
        Assert.Equal(4, methodNames.Count(name => name == saveName + "Async"));
        Assert.Equal(2, methodNames.Count(name => name == importName));
        Assert.Equal(2, methodNames.Count(name => name == resultName));

        foreach (Type receiverType in new[] { typeof(PdfDocument), typeof(PdfLogicalDocument) }) {
            MethodInfo[] receiverMethods = methods
                .Where(method => method.GetParameters()[0].ParameterType == receiverType)
                .ToArray();
            Assert.Equal(2, receiverMethods.Count(method => method.Name == saveName));
            Assert.Equal(2, receiverMethods.Count(method => method.Name == saveName + "Async"));
            Assert.Single(receiverMethods, method => method.Name == importName);
            Assert.Single(receiverMethods, method => method.Name == resultName);
            Assert.All(
                receiverMethods.Where(method => method.Name == saveName),
                method => Assert.NotEqual(typeof(void), method.ReturnType));
            Assert.All(
                receiverMethods.Where(method => method.Name == saveName + "Async"),
                method => Assert.True(
                    method.ReturnType.IsGenericType &&
                    method.ReturnType.GetGenericTypeDefinition() == typeof(Task<>),
                    method + " must return structured conversion evidence."));
        }
    }

    [Fact]
    public void FourPointZeroAdaptersDoNotExposeRetiredAmbiguousNames() {
        string[] retiredNames = [
            "ImportTablesToExcelDocument",
            "ImportTablesToExcelDocumentResult",
            "SaveTablesAsExcel",
            "SaveTablesAsExcelAsync",
            "ImportTablesToPowerPointPresentation",
            "ImportTablesToPowerPointPresentationResult",
            "SaveTablesAsPowerPoint",
            "SaveTablesAsPowerPointAsync"
        ];

        string[] publicNames = typeof(PdfExcelConverterExtensions)
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Concat(typeof(PowerPointPdfConverterExtensions).GetMethods(BindingFlags.Public | BindingFlags.Static))
            .Select(static method => method.Name)
            .ToArray();

        Assert.DoesNotContain(publicNames, retiredNames.Contains);
    }

    [Fact]
    public void FourPointZeroOptionNamesDescribeTheirOwningBridgeAndDirection() {
        Assert.Null(typeof(WordPdfSaveOptions).Assembly.GetType("OfficeIMO.Word.Pdf.PdfSaveOptions"));
        Assert.Null(typeof(PdfWordImportOptions).Assembly.GetType("OfficeIMO.Word.Pdf.PdfWordReadOptions"));
        Assert.Null(typeof(PdfExcelImportOptions).Assembly.GetType("OfficeIMO.Excel.Pdf.PdfExcelTableImportOptions"));
        Assert.Null(typeof(PdfPowerPointImportOptions).Assembly.GetType("OfficeIMO.PowerPoint.Pdf.PdfPowerPointTableImportOptions"));
        Assert.Null(typeof(PdfRtfImportOptions).Assembly.GetType("OfficeIMO.Rtf.Pdf.PdfRtfReadOptions"));
    }
}
