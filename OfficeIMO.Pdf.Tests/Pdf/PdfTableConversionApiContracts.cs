using System.Reflection;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfTableConversionApiContracts {
    [Theory]
    [InlineData(typeof(PdfExcelTableConverterExtensions), "SaveAsExcel", "ToExcelDocument", "ToExcelDocumentResult")]
    [InlineData(typeof(PdfWordConverterExtensions), "SaveAsWord", "ToWordDocument", "ToWordDocumentResult")]
    [InlineData(typeof(PowerPointPdfConverterExtensions), "SaveAsPowerPoint", "ToPowerPointPresentation", "ToPowerPointPresentationResult")]
    public void PdfToOfficeApisUseTheSameFluentShape(
        Type converterType,
        string saveName,
        string conversionName,
        string resultName) {
        MethodInfo[] methods = converterType
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .ToArray();
        string[] methodNames = methods.Select(method => method.Name).ToArray();

        Assert.Equal(2, methodNames.Count(name => name == saveName));
        Assert.Equal(2, methodNames.Count(name => name == saveName + "Async"));
        Assert.Single(methodNames, name => name == conversionName);
        Assert.Single(methodNames, name => name == resultName);
        Assert.All(
            methods.Where(method => method.Name == saveName || method.Name == saveName + "Async" || method.Name == conversionName || method.Name == resultName),
            method => Assert.Equal(typeof(PdfLogicalDocument), method.GetParameters()[0].ParameterType));
    }
}
