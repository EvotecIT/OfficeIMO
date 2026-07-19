using System.Reflection;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfTableConversionApiContracts {
    [Theory]
    [InlineData(typeof(PdfExcelTableConverterExtensions), "SaveTablesAsExcel", "ImportTablesToExcelDocument", "ImportTablesToExcelDocumentResult")]
    [InlineData(typeof(PowerPointPdfConverterExtensions), "SaveTablesAsPowerPoint", "ImportTablesToPowerPointPresentation", "ImportTablesToPowerPointPresentationResult")]
    public void TableOnlyAdaptersUseAnExplicitConsistentFluentShape(
        Type converterType,
        string saveName,
        string importName,
        string resultName) {
        MethodInfo[] methods = converterType
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .ToArray();
        string[] methodNames = methods.Select(method => method.Name).ToArray();

        Assert.Equal(2, methodNames.Count(name => name == saveName));
        Assert.Equal(2, methodNames.Count(name => name == saveName + "Async"));
        Assert.Single(methodNames, name => name == importName);
        Assert.Single(methodNames, name => name == resultName);
        Assert.All(
            methods.Where(method => method.Name == saveName || method.Name == saveName + "Async" || method.Name == importName || method.Name == resultName),
            method => Assert.Equal(typeof(PdfLogicalDocument), method.GetParameters()[0].ParameterType));
    }

    [Fact]
    public void TableOnlyAdaptersDoNotExposeFullDocumentConversionNames() {
        string[] misleadingNames = [
            "SaveAsExcel",
            "ToExcelDocument",
            "ToExcelDocumentResult",
            "SaveAsPowerPoint",
            "ToPowerPointPresentation",
            "ToPowerPointPresentationResult"
        ];

        string[] publicNames = typeof(PdfExcelTableConverterExtensions)
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Concat(typeof(PowerPointPdfConverterExtensions).GetMethods(BindingFlags.Public | BindingFlags.Static))
            .Select(static method => method.Name)
            .ToArray();

        Assert.DoesNotContain(publicNames, misleadingNames.Contains);
    }

    [Fact]
    public void WordAdapterKeepsFullDocumentConversionNames() {
        string[] publicNames = typeof(PdfWordConverterExtensions)
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Select(static method => method.Name)
            .ToArray();

        Assert.Equal(2, publicNames.Count(static name => name == "SaveAsWord"));
        Assert.Equal(2, publicNames.Count(static name => name == "SaveAsWordAsync"));
        Assert.Single(publicNames, static name => name == "ToWordDocument");
        Assert.Single(publicNames, static name => name == "ToWordDocumentResult");
    }
}
