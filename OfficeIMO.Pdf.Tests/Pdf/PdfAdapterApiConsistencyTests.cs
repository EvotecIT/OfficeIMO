using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using OfficeIMO.AsciiDoc.Pdf;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Latex.Pdf;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Mhtml;
using OfficeIMO.OneNote.Pdf;
using OfficeIMO.OpenDocument.Odp.Pdf;
using OfficeIMO.OpenDocument.Ods.Pdf;
using OfficeIMO.OpenDocument.Odt.Pdf;
using OfficeIMO.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Visio.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfAdapterApiConsistencyTests {
    public static IEnumerable<object[]> AdapterTypes() {
        yield return new object[] { typeof(AsciiDocPdfConverterExtensions) };
        yield return new object[] { typeof(WordPdfConverterExtensions) };
        yield return new object[] { typeof(ExcelPdfConverterExtensions) };
        yield return new object[] { typeof(PowerPointPdfConverterExtensions) };
        yield return new object[] { typeof(MarkdownPdfConverterExtensions) };
        yield return new object[] { typeof(RtfPdfConverterExtensions) };
        yield return new object[] { typeof(HtmlPdfConverterExtensions) };
        yield return new object[] { typeof(LatexPdfConverterExtensions) };
        yield return new object[] { typeof(MhtmlPdfConverterExtensions) };
        yield return new object[] { typeof(OneNoteSectionPdfConverterExtensions) };
        yield return new object[] { typeof(OneNoteNotebookPdfConverterExtensions) };
        yield return new object[] { typeof(OdtPdfConversionExtensions) };
        yield return new object[] { typeof(OdsPdfConversionExtensions) };
        yield return new object[] { typeof(OdpPdfConversionExtensions) };
        yield return new object[] { typeof(VisioPdfConverterExtensions) };
        yield return new object[] { typeof(OneNoteVisualPdfExtensions) };
    }

    public static IEnumerable<object[]> PdfOptionTypes() {
        yield return new object[] { typeof(AsciiDocToPdfOptions) };
        yield return new object[] { typeof(WordToPdfOptions) };
        yield return new object[] { typeof(ExcelToPdfOptions) };
        yield return new object[] { typeof(PowerPointToPdfOptions) };
        yield return new object[] { typeof(MarkdownToPdfOptions) };
        yield return new object[] { typeof(RtfToPdfOptions) };
        yield return new object[] { typeof(HtmlToPdfOptions) };
        yield return new object[] { typeof(LatexToPdfOptions) };
        yield return new object[] { typeof(OneNoteToPdfOptions) };
        yield return new object[] { typeof(VisioToPdfOptions) };
    }

    [Theory]
    [MemberData(nameof(AdapterTypes))]
    public void TypedPdfAdaptersExposeOneConsistentLifecyclePerSourceType(Type adapterType) {
        MethodInfo[] methods = adapterType.GetMethods(BindingFlags.Public | BindingFlags.Static);
        string target = adapterType == typeof(OneNoteVisualPdfExtensions) ? "VisualPdf" : "Pdf";
        Type[] sourceTypes = methods
            .Where(method => method.Name == $"To{target}Bytes")
            .Select(method => method.GetParameters()[0].ParameterType)
            .Distinct()
            .ToArray();

        Assert.NotEmpty(sourceTypes);
        foreach (Type sourceType in sourceTypes) {
            MethodInfo[] sourceMethods = methods
                .Where(method => method.GetParameters()[0].ParameterType == sourceType)
                .ToArray();

            Assert.Single(sourceMethods, method => method.Name == $"To{target}Bytes");
            Assert.Single(sourceMethods, method => method.Name == $"To{target}Document");
            Assert.Single(sourceMethods, method => method.Name == $"To{target}DocumentResult");
            Assert.Equal(2, sourceMethods.Count(method =>
                method.Name == $"SaveAs{target}" &&
                method.ReturnType == typeof(PdfSaveResult)));
            Assert.Equal(2, sourceMethods.Count(method =>
                method.Name == $"SaveAs{target}Result" &&
                method.ReturnType == typeof(PdfSaveResult)));
            Assert.Equal(2, sourceMethods.Count(method =>
                method.Name == $"SaveAs{target}Async" &&
                method.ReturnType == typeof(Task<PdfSaveResult>)));
            Assert.Equal(2, sourceMethods.Count(method =>
                method.Name == $"SaveAs{target}ResultAsync" &&
                method.ReturnType == typeof(Task<PdfSaveResult>)));

            string[] asynchronousConversionMethods = [$"To{target}BytesAsync", $"To{target}DocumentAsync", $"To{target}DocumentResultAsync"];
            int asynchronousConversionMethodCount = sourceMethods.Count(method => asynchronousConversionMethods.Contains(method.Name));
            Assert.True(
                asynchronousConversionMethodCount is 0 or 3,
                $"{adapterType.Name} must expose either the complete asynchronous conversion trio for an asynchronous engine or none for a synchronous engine.");
            foreach (string methodName in asynchronousConversionMethods.Where(_ => asynchronousConversionMethodCount > 0)) {
                Assert.Single(sourceMethods, method => method.Name == methodName);
            }
        }

        Assert.DoesNotContain(methods, method =>
            method.Name == $"To{target}" || method.Name == $"To{target}Async" ||
            method.Name == $"TrySaveAs{target}" || method.Name == $"TrySaveAs{target}Async");
        Assert.DoesNotContain(methods, method => method.GetParameters()[0].ParameterType == typeof(string));
        Assert.All(methods, method => {
            Assert.Equal(typeof(CancellationToken), method.GetParameters().Last().ParameterType);
            Assert.True(method.GetParameters().Last().IsOptional, method.ToString());
        });
        Assert.All(
            methods.SelectMany(static method => method.GetParameters())
                .Where(static parameter => parameter.ParameterType == typeof(CancellationToken)),
            static parameter => {
                Assert.Equal("cancellationToken", parameter.Name);
                Assert.Equal(((MethodBase)parameter.Member).GetParameters().Length - 1, parameter.Position);
            });
    }

    [Theory]
    [MemberData(nameof(PdfOptionTypes))]
    public void PdfExportOptionsUseTargetNamingAndDoNotCaptureOperationCancellation(Type optionsType) {
        Assert.EndsWith("ToPdfOptions", optionsType.Name, StringComparison.Ordinal);
        Assert.DoesNotContain(
            optionsType.GetProperties(BindingFlags.Public | BindingFlags.Instance),
            static property => property.PropertyType == typeof(CancellationToken));
    }
}
