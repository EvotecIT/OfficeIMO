using System.Reflection;
using OfficeIMO.Excel;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.PowerPoint;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfConversionApiConsistencyTests {
    public static TheoryData<Type, Type> PrimaryConverters => new() {
        { typeof(WordPdfConverterExtensions), typeof(WordDocument) },
        { typeof(ExcelPdfConverterExtensions), typeof(ExcelDocument) },
        { typeof(MarkdownPdfConverterExtensions), typeof(MarkdownDoc) },
        { typeof(HtmlPdfConverterExtensions), typeof(HtmlConversionDocument) },
        { typeof(PowerPointPdfConverterExtensions), typeof(PowerPointPresentation) },
        { typeof(RtfPdfConverterExtensions), typeof(RtfDocument) }
    };

    [Theory]
    [MemberData(nameof(PrimaryConverters))]
    public void PrimaryConverter_ExposesCanonicalDocumentAndSaveFamilies(Type extensionType, Type sourceType) {
        MethodInfo[] methods = extensionType.GetMethods(BindingFlags.Public | BindingFlags.Static);

        AssertHasSourceMethod(methods, sourceType, "ToPdfDocument");
        AssertHasSourceMethod(methods, sourceType, "ToPdfDocumentResult");
        AssertHasSourceMethod(methods, sourceType, "ToPdfBytes");
        AssertHasDestinationOverloads(methods, sourceType, "SaveAsPdf");
        AssertHasDestinationOverloads(methods, sourceType, "SaveAsPdfResult");
        AssertHasDestinationOverloads(methods, sourceType, "SaveAsPdfAsync");
        AssertHasDestinationOverloads(methods, sourceType, "SaveAsPdfResultAsync");
    }

    private static void AssertHasSourceMethod(MethodInfo[] methods, Type sourceType, string name) {
        Assert.Contains(methods, method =>
            method.Name == name &&
            method.GetParameters() is { Length: > 0 } parameters &&
            parameters[0].ParameterType == sourceType);
    }

    private static void AssertHasDestinationOverloads(MethodInfo[] methods, Type sourceType, string name) {
        Assert.Contains(methods, method => HasDestination(method, sourceType, name, typeof(string)));
        Assert.Contains(methods, method => HasDestination(method, sourceType, name, typeof(Stream)));
    }

    private static bool HasDestination(MethodInfo method, Type sourceType, string name, Type destinationType) {
        ParameterInfo[] parameters = method.GetParameters();
        return method.Name == name &&
               parameters.Length >= 2 &&
               parameters[0].ParameterType == sourceType &&
               parameters[1].ParameterType == destinationType;
    }
}
