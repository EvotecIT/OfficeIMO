using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfMergeResourceRegressionTests {
    [Fact]
    public void Append_Preserves_File_Namespace_And_Metadata_Tables_And_Reports_Conflicting_Global_Values() {
        RtfDocument destination = RtfDocument.Create();
        destination.AddFileReference("file:///destination.docx");
        destination.AddXmlNamespace(0, "urn:destination");
        destination.Info.Title = "Destination";
        destination.AddDocumentVariable("Shared", "Destination");
        RtfDocument source = RtfDocument.Create();
        source.AddParagraph("Imported");
        source.AddFileReference("file:///source.docx").RelativePathStart = 8;
        source.AddXmlNamespace(0, "urn:source");
        source.Info.Title = "Source";
        source.Info.Author = "Author";
        source.AddDocumentVariable("Shared", "Source");
        source.AddUserProperty("Unique", RtfUserProperty.TextType, "Value");
        source.AddRevisionSaveId(123);

        RtfDocumentMergeResult result = destination.AppendDocument(source);
        Assert.Equal(2, destination.FileReferences.Count);
        Assert.Equal(2, destination.XmlNamespaces.Count);
        Assert.Equal(2, destination.FileReferences.Select(item => item.Id).Distinct().Count());
        Assert.Equal(2, destination.XmlNamespaces.Select(item => item.Id).Distinct().Count());
        Assert.Equal(8, destination.FileReferences[1].RelativePathStart);
        Assert.Equal("Destination", destination.Info.Title);
        Assert.Equal("Author", destination.Info.Author);
        Assert.Equal("Value", Assert.Single(destination.UserProperties).StaticValue);
        Assert.Contains(123, destination.RevisionSaveIds);
        Assert.Equal("Destination", Assert.Single(destination.DocumentVariables).Value);
        Assert.Equal(new[] { "DocumentVariables/Shared", "Info/Title" }, result.Report.Diagnostics.Where(item => item.Code == "RtfMergeMetadataConflict").Select(item => item.SourcePath));
        Assert.Throws<RtfConversionLossException>(() => result.Report.RequireNoLoss());
        RtfDocument reopened = RtfDocument.Read(destination.ToRtf(), new RtfReadOptions { ReadFileReferences = true }).Document;
        Assert.Equal("file:///source.docx", reopened.FileReferences[1].Path);
        Assert.Equal("urn:source", reopened.XmlNamespaces[1].Uri);
        Assert.Equal(0, source.FileReferences[0].Id);
        Assert.Equal(0, source.XmlNamespaces[0].Id);
    }

    [Fact]
    public void Append_Imports_Kind_Specific_Style_Chains_Defaults_And_List_Counters_Without_Resource_Collisions() {
        RtfDocument destination = RtfDocument.Create();
        destination.AddStyle(0, "Destination Normal").Bold = true;
        destination.AddStyle(5, "Destination style");
        destination.AddColor(0, 0, 255);
        destination.AddListDefinition(100).AddLevel().StartAt = 90;
        destination.AddListOverride(3, 100);
        destination.AddParagraph("Destination").SetList(3);

        RtfDocument source = RtfDocument.Create();
        int font = source.AddFont("Consolas");
        source.Settings.DefaultFontId = font;
        source.Settings.DefaultLanguageId = 1045;
        int red = source.AddColor(255, 0, 0);
        source.AddStyle(0, "Source Normal").Italic = true;
        RtfStyle parent = source.AddStyle(5, "Source base");
        parent.ForegroundColorIndex = red;
        parent.Bold = true;
        parent.BasedOnStyleId = 0;
        source.AddStyle(6, "Source child").BasedOnStyleId = 5;
        RtfStyle character = source.AddStyle(5, "Source character", RtfStyleKind.Character);
        character.FontSize = 19;
        character.Italic = false;
        RtfListLevel level = source.AddListDefinition(100).AddLevel();
        level.StartAt = 4;
        level.NumberFormat = 1;
        source.AddListOverride(3, 100);
        RtfParagraph paragraph = source.AddParagraph("Imported").SetList(3);
        paragraph.StyleId = 6;
        paragraph.Runs[0].StyleId = 5;
        source.AddParagraph("Source default");

        RtfDocumentMergeResult result = destination.AppendDocument(source);
        result.Report.RequireNoLoss();
        RtfParagraph imported = destination.Paragraphs[1];
        Assert.NotEqual(6, imported.StyleId);
        Assert.NotEqual(3, imported.ListId);
        Assert.NotEqual(imported.StyleId, imported.Runs[0].StyleId);
        RtfRun effective = destination.ResolveRunFormatting(imported, imported.Runs[0]);
        Assert.True(effective.Bold);
        Assert.False(effective.Italic);
        Assert.Equal(19, effective.FontSize);
        Assert.Equal(1045, effective.LanguageId);
        Assert.Equal("Consolas", destination.Fonts.Single(item => item.Id == effective.FontId).Name);
        Assert.Equal((byte)255, destination.GetColor(effective.ForegroundColorIndex!.Value)!.Red);
        Assert.True(destination.ResolveRunFormatting(destination.Paragraphs[2], destination.Paragraphs[2].Runs[0]).Italic);
        Assert.False(destination.ResolveRunFormatting(destination.Paragraphs[2], destination.Paragraphs[2].Runs[0]).Bold);
        foreach (RtfDocument value in new[] { destination, RtfDocument.Read(destination.ToRtf()).Document }) {
            var numbering = new RtfListNumbering(value);
            Assert.Equal("90.", numbering.Next(value.Paragraphs[0])!.Text);
            Assert.Equal("IV.", numbering.Next(value.Paragraphs[1])!.Text);
        }
        Assert.Equal(6, paragraph.StyleId);
        Assert.Equal(5, paragraph.Runs[0].StyleId);
        destination.Styles.Single(item => item.Name == "Source character").FontSize = 21;
        Assert.Equal(19, character.FontSize);
    }

    [Fact]
    public void Append_Remaps_Shared_Content_Exactly_Once_And_Does_Not_Deduplicate_Distinct_Font_Metadata() {
        RtfDocument destination = RtfDocument.Create();
        destination.AddColor(0, 255, 0);
        RtfDocument source = RtfDocument.Create();
        int red = source.AddColor(255, 0, 0);
        source.Fonts[0].Embedding = new RtfFontEmbedding { Data = new byte[] { 1, 2, 3 }, FileName = "font.bin" };
        source.Fonts[0].AlternateName = "Embedded family";
        RtfParagraph paragraph = source.AddObject().Result;
        paragraph.AddText("Repeated");
        paragraph.Runs[0].ForegroundColorIndex = red;
        source.AddBlock(paragraph);

        destination.AppendDocument(source).Report.RequireNoLoss();
        Assert.Same(Assert.IsType<RtfObject>(destination.Blocks[0]).Result, destination.Blocks[1]);
        Assert.Equal(2, destination.Paragraphs[0].Runs[0].ForegroundColorIndex);
        Assert.Equal(2, destination.Fonts.Count);
        RtfFont imported = destination.Fonts.Single(item => item.Embedding != null);
        Assert.Equal("Embedded family", imported.AlternateName);
        imported.Embedding!.Data[0] = 9;
        Assert.Equal(1, source.Fonts[0].Embedding!.Data[0]);
    }
}
