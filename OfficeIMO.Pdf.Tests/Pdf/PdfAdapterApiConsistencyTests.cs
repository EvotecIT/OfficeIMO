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
    public void TypedPdfAdaptersExposeOneConsistentLifecyclePerSourceType(Type adapterType) {
        MethodInfo[] methods = adapterType.GetMethods(BindingFlags.Public | BindingFlags.Static);
        Type[] sourceTypes = methods
            .Where(method => method.Name == "ToPdf")
            .Select(method => method.GetParameters()[0].ParameterType)
            .Distinct()
            .ToArray();

        Assert.NotEmpty(sourceTypes);
        foreach (Type sourceType in sourceTypes) {
            MethodInfo[] sourceMethods = methods
                .Where(method => method.GetParameters()[0].ParameterType == sourceType)
                .ToArray();

            Assert.Single(sourceMethods, method => method.Name == "ToPdf");
            Assert.Single(sourceMethods, method => method.Name == "ToPdfDocument");
            Assert.Single(sourceMethods, method => method.Name == "ToPdfDocumentResult");
            Assert.Equal(2, sourceMethods.Count(method => method.Name == "SaveAsPdf"));
            Assert.Equal(2, sourceMethods.Count(method => method.Name == "TrySaveAsPdf"));
            Assert.Equal(2, sourceMethods.Count(method => method.Name == "SaveAsPdfAsync"));
            Assert.Equal(2, sourceMethods.Count(method => method.Name == "TrySaveAsPdfAsync"));

            string[] asynchronousConversionMethods = ["ToPdfAsync", "ToPdfDocumentAsync", "ToPdfDocumentResultAsync"];
            int asynchronousConversionMethodCount = sourceMethods.Count(method => asynchronousConversionMethods.Contains(method.Name));
            Assert.True(
                asynchronousConversionMethodCount is 0 or 3,
                $"{adapterType.Name} must expose either the complete asynchronous conversion trio for an asynchronous engine or none for a synchronous engine.");
            foreach (string methodName in asynchronousConversionMethods.Where(_ => asynchronousConversionMethodCount > 0)) {
                Assert.Single(sourceMethods, method => method.Name == methodName);
            }
        }

        Assert.DoesNotContain(methods, method => method.GetParameters()[0].ParameterType == typeof(string));
    }
}
