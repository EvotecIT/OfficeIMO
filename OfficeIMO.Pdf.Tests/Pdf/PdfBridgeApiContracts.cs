using System.Reflection;
using System.Threading.Tasks;
using OfficeIMO.Excel.Pdf;
using OfficeIMO.Html.Pdf;
using OfficeIMO.PowerPoint.Pdf;
using OfficeIMO.Reader.Pdf;
using OfficeIMO.Rtf.Pdf;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Pdf.Tests;

public sealed class PdfBridgeApiContracts {
    [Theory]
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

        foreach (Type receiverType in new[] { typeof(PdfDocument), typeof(PdfDocumentReadResult) }) {
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
    public void ExcelTableRecoveryUsesTheSameNarrowFacadeOnOpenedAndLogicalPdfDocuments() {
        MethodInfo[] methods = typeof(PdfExcelTableConverterExtensions)
            .GetMethods(BindingFlags.Public | BindingFlags.Static);
        string[] methodNames = methods.Select(static method => method.Name).ToArray();

        Assert.Equal(4, methodNames.Count(static name => name == "SaveTablesAsExcel"));
        Assert.Equal(4, methodNames.Count(static name => name == "SaveTablesAsExcelAsync"));
        Assert.Equal(2, methodNames.Count(static name => name == "ImportTablesToExcelDocument"));
        Assert.Equal(2, methodNames.Count(static name => name == "ImportTablesToExcelDocumentResult"));

        foreach (Type receiverType in new[] { typeof(PdfDocument), typeof(PdfDocumentReadResult) }) {
            MethodInfo[] receiverMethods = methods
                .Where(method => method.GetParameters()[0].ParameterType == receiverType)
                .ToArray();

            Assert.Equal(2, receiverMethods.Count(static method => method.Name == "SaveTablesAsExcel"));
            Assert.Equal(2, receiverMethods.Count(static method => method.Name == "SaveTablesAsExcelAsync"));
            Assert.Single(receiverMethods, static method => method.Name == "ImportTablesToExcelDocument");
            Assert.Single(receiverMethods, static method => method.Name == "ImportTablesToExcelDocumentResult");
        }
    }

    [Fact]
    public void GeneralAndNarrowRoutesDoNotCompeteForTheSameFacadeNames() {
        string[] retiredPowerPointTableNames = [
            "ImportTablesToPowerPointPresentation",
            "ImportTablesToPowerPointPresentationResult",
            "SaveTablesAsPowerPoint",
            "SaveTablesAsPowerPointAsync"
        ];

        string[] excelNames = typeof(PdfExcelTableConverterExtensions)
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Select(static method => method.Name)
            .ToArray();
        string[] powerPointNames = typeof(PowerPointPdfConverterExtensions)
            .GetMethods(BindingFlags.Public | BindingFlags.Static)
            .Select(static method => method.Name)
            .ToArray();

        Assert.DoesNotContain("ToExcelDocument", excelNames);
        Assert.DoesNotContain("ToExcelDocumentResult", excelNames);
        Assert.DoesNotContain("SaveAsExcel", excelNames);
        Assert.DoesNotContain("SaveAsExcelAsync", excelNames);
        Assert.DoesNotContain(powerPointNames, retiredPowerPointTableNames.Contains);
    }

    [Fact]
    public void FourPointZeroOptionNamesDescribeTheirOwningBridgeAndDirection() {
        Assert.Null(typeof(WordToPdfOptions).Assembly.GetType("OfficeIMO.Word.Pdf.PdfSaveOptions"));
        Assert.Null(typeof(PdfToWordOptions).Assembly.GetType("OfficeIMO.Word.Pdf.PdfWordReadOptions"));
        Assert.Null(typeof(PdfTablesToExcelOptions).Assembly.GetType("OfficeIMO.Excel.Pdf.PdfExcelImportOptions"));
        Assert.Null(typeof(PdfToPowerPointOptions).Assembly.GetType("OfficeIMO.PowerPoint.Pdf.PdfPowerPointTableImportOptions"));
        Assert.Null(typeof(PdfToRtfOptions).Assembly.GetType("OfficeIMO.Rtf.Pdf.PdfRtfReadOptions"));
    }

    [Fact]
    public void OpenedDocumentAdaptersExposeTheCanonicalSemanticReadOptions() {
        foreach (Type optionType in new[] {
            typeof(PdfToWordOptions),
            typeof(PdfTablesToExcelOptions),
            typeof(PdfToRtfOptions),
            typeof(PdfToHtmlOptions),
            typeof(PdfToPowerPointOptions),
            typeof(ReaderPdfOptions)
        }) {
            PropertyInfo property = Assert.IsAssignableFrom<PropertyInfo>(optionType.GetProperty("ReadOptions"));
            Assert.Equal(typeof(PdfReadOptions), property.PropertyType);
            Assert.True(property.CanRead);
            Assert.True(property.CanWrite);
        }
    }

    [Fact]
    public void OpenedDocumentAdaptersPropagateTheCanonicalSemanticPageLimit() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("First page"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second page"))
            .ToBytes();
        PdfDocument document = PdfDocument.Load(pdf);
        var readOptions = new PdfReadOptions {
            Pipeline = new PdfUnderstandingPipelineOptions { MaxPages = 1 }
        };

        Assert.Throws<PdfReadLimitException>(() =>
            document.ToWordDocument(new PdfToWordOptions { ReadOptions = readOptions }));
        Assert.Throws<PdfReadLimitException>(() =>
            document.ImportTablesToExcelDocument(new PdfTablesToExcelOptions { ReadOptions = readOptions }));
        Assert.Throws<PdfReadLimitException>(() =>
            document.ToRtfDocument(new PdfToRtfOptions { ReadOptions = readOptions }));
        Assert.Throws<PdfReadLimitException>(() =>
            document.ToHtml(new PdfToHtmlOptions { ReadOptions = readOptions }));
        Assert.Throws<PdfReadLimitException>(() =>
            document.ToPowerPointPresentation(new PdfToPowerPointOptions { ReadOptions = readOptions }));
        using var stream = new MemoryStream(pdf, writable: false);
        Assert.Throws<PdfReadLimitException>(() =>
            PdfReaderAdapter.ReadDocument(stream, pdfOptions: new ReaderPdfOptions { ReadOptions = readOptions }));
    }

    [Fact]
    public void WordImportUsesCanonicalPageSelectionWithinTheConfiguredLimit() {
        byte[] pdf = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("First page marker"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second page marker"))
            .ToBytes();
        PdfDocument document = PdfDocument.Load(pdf);
        var options = new PdfToWordOptions {
            ReadOptions = new PdfReadOptions {
                PageSelection = PdfPageSelection.From(2),
                Pipeline = new PdfUnderstandingPipelineOptions { MaxPages = 1 }
            }
        };

        using OfficeIMO.Word.WordDocument word = document.ToWordDocument(options);
        string text = string.Join(" ", word.Paragraphs.Select(static paragraph => paragraph.Text));

        Assert.Contains("Second page marker", text, StringComparison.Ordinal);
        Assert.DoesNotContain("First page marker", text, StringComparison.Ordinal);
        Assert.Same(options.ReadOptions, options.Clone().ReadOptions);
    }

    [Fact]
    public void GeneralPowerPointRouteUsesConversionResultsAndExplicitTableEntries() {
        Assembly assembly = typeof(PdfPowerPointConversionResult).Assembly;

        Assert.Null(assembly.GetType("OfficeIMO.PowerPoint.Pdf.PdfPowerPointImportResult"));
        Assert.Null(assembly.GetType("OfficeIMO.PowerPoint.Pdf.PdfPowerPointImportReport"));
        Assert.NotNull(typeof(PdfPowerPointConversionReport).GetProperty("TableEntries"));
        Assert.Null(typeof(PdfPowerPointConversionReport).GetProperty("Entries"));
    }
}
