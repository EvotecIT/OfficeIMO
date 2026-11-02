using OfficeIMO.Excel;
using OfficeIMO.PowerPoint;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
using OfficeIMO.Reader.Excel;
using OfficeIMO.Reader.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.Tests;

public sealed class OfficeDocumentSummaryRoutingTests {
    [Fact]
    public void RegisteredDetectionDefaultsAreNotReplacedByInspectionDefaults() {
        using var word = WordDocument.Create();
        word.AddParagraph("Detection policy");
        using var input = new MemoryStream(word.ToBytes());
        var defaults = new ReaderOptions {
            DetectionMode = ReaderDetectionMode.ExtensionOnly,
            DetectionInspectContainers = false,
            DetectionMaxProbeBytes = 17,
            DetectionMaxContainerEntries = 2
        };
        OfficeDocumentReader reader = new OfficeDocumentReaderBuilder().AddAllOfficeIMOHandlers()
            .WithDefaultOptions(defaults).Build();
        var options = new OfficeInspectOptions { IncludeSummary = true, IncludeContent = false,
            Reader = reader.DefaultOptions, DocumentReader = reader };
        OfficeDocumentInfo info = OfficeDocument.Inspect(input, "source.txt", options);
        Assert.False(info.Detection!.ContentInspected);
        Assert.Equal(0, info.Detection.InspectedBytes);
        Assert.Equal(ReaderDetectionMode.ExtensionOnly, defaults.DetectionMode);
    }

    [Theory]
    [InlineData("Excel")]
    [InlineData("PowerPoint")]
    public void NormalizedContainersDoNotEstablishTotalSourcePages(string format) {
        OfficeDocumentReadResult read;
        if (format == "Excel") {
            using var workbook = ExcelDocument.Create();
            workbook.AddSheet("Data").CellValue(1, 1, "Inventory");
            workbook.AddSheet("Empty");
            read = workbook.ReadDocument();
        } else {
            using var presentation = PowerPointDocument.Create();
            presentation.AddSlide().AddTextBox("Inventory");
            presentation.AddSlide();
            read = presentation.ReadDocument();
        }
        Assert.NotEmpty(read.Pages);
        OfficeDocumentInfo info = OfficeDocument.Inspect((object)read,
            new OfficeInspectOptions { IncludeContent = false });
        Assert.Null(info.PageCount);
        Assert.Null(info.SheetCount);
        Assert.Null(info.SlideCount);
    }

    [Theory]
    [InlineData("docx")]
    [InlineData("xlsx")]
    [InlineData("pptx")]
    public void OneBatchPasswordDoesNotRejectUnencryptedOfficeSources(string extension) {
        string root = Directory.CreateTempSubdirectory("officeimo-password-inventory-").FullName;
        try {
            string path = Path.Combine(root, "plain." + extension);
            if (extension == "docx") {
                using var word = WordDocument.Create(); word.AddParagraph("Plain source"); word.Save(path);
            } else if (extension == "xlsx") {
                using var excel = ExcelDocument.Create(); excel.AddSheet("Data"); excel.Save(path);
            } else {
                using var presentation = PowerPointDocument.Create(); presentation.AddSlide(); presentation.Save(path);
            }
            byte[] original = File.ReadAllBytes(path);
            var options = new OfficeInspectOptions { IncludeSummary = true, IncludeContent = false };
            options.Reader.OpenPassword = "Synthetic-batch-password";
            OfficeDocumentInfo info = OfficeDocument.Inspect(path, options);
            Assert.False(info.IsEncrypted);
            Assert.Equal(original, File.ReadAllBytes(path));
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData("vdx")]
    [InlineData("vsx")]
    [InlineData("vtx")]
    public void LegacyVisioXmlUsesTheGenerationAwareNativeOwner(string extension) {
        byte[] bytes = Encoding.UTF8.GetBytes("""
            <VisioDocument xmlns="http://schemas.microsoft.com/visio/2003/core">
              <Pages><Page ID="0" Name="Inventory"><Shapes><Shape ID="1"><Text>Legacy inventory</Text></Shape></Shapes></Page></Pages>
            </VisioDocument>
            """);
        using var input = new MemoryStream(bytes);
        OfficeDocumentInfo info = OfficeDocument.Inspect(input, "source." + extension,
            new OfficeInspectOptions { IncludeSummary = true, IncludeContent = false, Detailed = true });
        Assert.Equal(ReaderInputKind.Visio, info.Format);
        Assert.Equal(1, info.PageCount);
        Assert.IsType<OfficeIMO.Visio.VisioInspectionSnapshot>(info.Details);
    }

    [Theory]
    [InlineData("vsd")]
    [InlineData("vss")]
    [InlineData("vst")]
    public void IndependentLegacyVisioBinaryProfilesRetainNativeInventory(string extension) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Fixtures", "Visio", "pronom-visio2003." + extension));
        using var input = new MemoryStream(bytes);
        OfficeDocumentInfo info = OfficeDocument.Inspect(input, "source." + extension,
            new OfficeInspectOptions { IncludeSummary = true, IncludeContent = false, Detailed = true });
        Assert.Equal(ReaderInputKind.Visio, info.Format);
        Assert.NotNull(info.PageCount);
        Assert.IsType<OfficeIMO.Visio.VisioInspectionSnapshot>(info.Details);
        Assert.Equal(0, input.Position);
        Assert.Equal(bytes, input.ToArray());
    }
}
