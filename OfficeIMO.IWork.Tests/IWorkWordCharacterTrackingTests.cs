using OfficeIMO.IWork;
using OfficeIMO.Word;
using OfficeIMO.Word.IWork;
using Xunit;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    [Theory]
    [InlineData(false, -0.25f, -200)]
    [InlineData(false, 0f, 0)]
    [InlineData(false, 0.25f, 200)]
    [InlineData(true, -0.25f, -200)]
    [InlineData(true, 0f, 0)]
    [InlineData(true, 0.25f, 200)]
    public void Pages_character_tracking_survives_saved_body_and_table_runs(bool tableCell, float tracking, int twips) {
        byte[] attributes = AttributeTable(8, AttributeEntry(0, ReferenceField(2, 19)));
        byte[] style = ArchiveRecord(19, 2021, BytesField(11, Message(FloatField(3, 40), FloatField(27, tracking))));
        using var package = tableCell
            ? SelectedRichTextPackage(IWorkDocumentKind.Pages, attributes, additionalRecords: style)
            : CreatePackage(("Index/Document.iwa", FrameIwa(Message(
                ArchiveRecord(1, 10000, ReferenceField(4, 2), new ulong[] { 2 }),
                ArchiveRecord(2, 2001, Message(StringField(3, "Body"), attributes)), style))));
        using var converted = WordIWorkConverter.ConvertPagesToWordResult(package,
            conversionOptions: new IWorkConversionOptions { Mode = IWorkConversionMode.EditableOnly,
                AllowPartialEditableReconstruction = tableCell });
        if (!tableCell) converted.Report.RequireCompleteEditableReconstruction();
        using var bytes = new MemoryStream();
        converted.Value.Save(bytes);
        bytes.Position = 0;
        using WordDocument reopened = WordDocument.Load(bytes);
        WordParagraph run = tableCell ? reopened.Tables[0].Rows[0].Cells[0].Paragraphs.First(paragraph => paragraph.Text == "Value")
            : reopened.Paragraphs.First(paragraph => paragraph.Text == "Body");
        Assert.Equal(twips, run.Spacing);
    }
}
