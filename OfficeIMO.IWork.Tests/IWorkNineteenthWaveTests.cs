using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Excel;
using OfficeIMO.IWork;
using OfficeIMO.Internal;
using OfficeIMO.PowerPoint;
using OfficeIMO.Word;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Fact]
    public void Reused_pages_headers_charge_the_character_budget_for_each_destination() {
        using MemoryStream package = CreatePagesPackageWithSharedHeader();
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, IWorkDocumentKind.Pages,
            new IWorkReadOptions { MaximumProjectedTextCharacters = 12 });

        Assert.Throws<InvalidDataException>(() => source.ReadPages());
    }

    [Fact]
    public void Keynote_right_paragraph_indents_use_visual_fallback() {
        using MemoryStream package = CreateKeynotePackageWithRightIndent();

        using var result = PowerPointIWorkConverter.ConvertKeynoteToPowerPointResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
    }

    [Theory]
    [InlineData("This Numbers sheet name is longer than thirty one characters")]
    [InlineData("Invalid/Sheet")]
    [InlineData("'Trimmed'")]
    public void Numbers_sheet_names_that_xlsx_would_change_use_visual_fallback(string sheetName) {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("Table", 1, 1, 42d)
        }, includePreview: true, sheetNameBytes: Encoding.UTF8.GetBytes(sheetName));

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
    }

    [Fact]
    public void Case_insensitive_destination_sheet_name_collisions_use_visual_fallback() {
        using MemoryStream package = CreateNumbersPackage(new[] {
            new TableSpec("A", 1, 1, 1d),
            new TableSpec("a", 1, 1, 2d)
        }, includePreview: true);

        using var result = ExcelIWorkConverter.ConvertNumbersToExcelResult(package);

        Assert.True(result.IsVisualFallback);
        Assert.True(result.Projection.HasEditableContent);
    }

    [Fact]
    public void Files_opened_below_a_filesystem_root_pass_physical_root_containment() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-iwork-root-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "entry.iwa");
        File.WriteAllBytes(path, new byte[] { 1, 2, 3 });
        try {
            string root = Path.GetPathRoot(Path.GetFullPath(path))
                ?? throw new InvalidDataException("The test path has no filesystem root.");

            using FileStream stream = OfficePathIdentity.OpenRegularFileForRead(path, root, 4096);

            Assert.Equal(3, stream.Length);
        } finally {
            File.Delete(path);
            Directory.Delete(directory);
        }
    }

    private static MemoryStream CreatePagesPackageWithSharedHeader() {
        const ulong documentId = 1;
        const ulong bodyId = 2;
        const ulong firstSectionId = 3;
        const ulong secondSectionId = 4;
        const ulong firstHeaderFooterId = 5;
        const ulong secondHeaderFooterId = 6;
        const ulong headerId = 7;
        byte[] sectionTable = Message(
            BytesField(1, Message(ReferenceField(2, firstSectionId))),
            BytesField(1, Message(ReferenceField(2, secondSectionId))));
        byte[] records = Message(
            ArchiveRecord(documentId, 10000, Message(ReferenceField(4, bodyId)), new[] { bodyId }),
            ArchiveRecord(bodyId, 2001,
                Message(StringField(3, "A\u0004B"), BytesField(17, sectionTable)),
                new[] { firstSectionId, secondSectionId }),
            ArchiveRecord(firstSectionId, 10011, Message(ReferenceField(25, firstHeaderFooterId)),
                new[] { firstHeaderFooterId }),
            ArchiveRecord(secondSectionId, 10011, Message(ReferenceField(25, secondHeaderFooterId)),
                new[] { secondHeaderFooterId }),
            ArchiveRecord(firstHeaderFooterId, 10143, Message(ReferenceField(1, headerId)),
                new[] { headerId }),
            ArchiveRecord(secondHeaderFooterId, 10143, Message(ReferenceField(1, headerId)),
                new[] { headerId }),
            ArchiveRecord(headerId, 2001, Message(StringField(3, "Header"))));
        return CreatePackage(("Index/Document.iwa", FrameIwa(records)));
    }

    private static MemoryStream CreateKeynotePackageWithRightIndent() {
        const ulong documentId = 1;
        const ulong showId = 2;
        const ulong nodeId = 3;
        const ulong slideId = 4;
        const ulong shapeId = 5;
        const ulong storageId = 6;
        const ulong styleId = 7;
        byte[] styleTable = Message(BytesField(1,
            Message(VarintField(1, 0), ReferenceField(2, styleId))));
        byte[] records = Message(
            ArchiveRecord(documentId, 1, Message(ReferenceField(2, showId))),
            ArchiveRecord(showId, 2, KeynoteShow(Message(ReferenceField(2, nodeId)))),
            ArchiveRecord(nodeId, 4, Message(ReferenceField(2, slideId))),
            ArchiveRecord(slideId, 5, Message(ReferenceField(5, shapeId))),
            ArchiveRecord(shapeId, 2011, Message(ReferenceField(2, storageId))),
            ArchiveRecord(storageId, 2001,
                Message(StringField(3, "Indented"), BytesField(5, styleTable)), new[] { styleId }),
            ArchiveRecord(styleId, 2022, Message(BytesField(12, Message(FloatField(19, 12f))))));
        return CreatePackage(
            ("Index/Slide.iwa", FrameIwa(records)),
            ("preview.png", ValidPreviewPng()));
    }
}
