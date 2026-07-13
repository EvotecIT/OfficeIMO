using System.Reflection;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfAdapterApiConsistencyTests {
    public static IEnumerable<object[]> AdapterTypes() {
        yield return new object[] { typeof(WordPdfConverterExtensions) };
        yield return new object[] { typeof(ExcelPdfConverterExtensions) };
        yield return new object[] { typeof(PowerPointPdfConverterExtensions) };
        yield return new object[] { typeof(MarkdownPdfConverterExtensions) };
        yield return new object[] { typeof(RtfPdfConverterExtensions) };
        yield return new object[] { typeof(HtmlPdfConverterExtensions) };
    }

    [Theory]
    [MemberData(nameof(AdapterTypes))]
    public void TypedPdfAdaptersExposeOneConsistentLifecycle(Type adapterType) {
        MethodInfo[] methods = adapterType.GetMethods(BindingFlags.Public | BindingFlags.Static);

        Assert.Single(methods, method => method.Name == "ToPdf");
        Assert.Single(methods, method => method.Name == "ToPdfDocument");
        Assert.Single(methods, method => method.Name == "ToPdfDocumentResult");
        Assert.Equal(2, methods.Count(method => method.Name == "SaveAsPdf"));
        Assert.Equal(2, methods.Count(method => method.Name == "TrySaveAsPdf"));
        Assert.Equal(2, methods.Count(method => method.Name == "SaveAsPdfAsync"));
        Assert.Equal(2, methods.Count(method => method.Name == "TrySaveAsPdfAsync"));
        Assert.DoesNotContain(methods, method => method.GetParameters()[0].ParameterType == typeof(string));
    }
}
