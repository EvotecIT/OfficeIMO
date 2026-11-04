using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class NativeWordReviewRemediationTests {
    [Fact]
    public void SectionFindResolvesAllInheritedStoriesAndStopsAtExplicitEmptyParts() {
        using var source = WordDocument.Create();
        source.AddParagraph("needle outside");
        WordHeaderFooter[] stories = { source.PageSetup.Header, source.PageSetup.FirstPageHeader, source.PageSetup.EvenPageHeader,
            source.PageSetup.Footer, source.PageSetup.FirstPageFooter, source.PageSetup.EvenPageFooter };
        for (int index = 0; index < stories.Length; index++) {
            stories[index].AddParagraph("needle story " + index);
            stories[index].AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "needle story table " + index;
        }
        var inherited = source.AddSection();
        inherited.AddParagraph("needle selected body");
        inherited.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "needle selected table";
        inherited._sectionProperties.RemoveAllChildren<HeaderReference>();
        inherited._sectionProperties.RemoveAllChildren<FooterReference>();
        var empty = source.AddSection();
        _ = empty.PageSetup.Header; _ = empty.PageSetup.FirstPageHeader; _ = empty.PageSetup.EvenPageHeader;
        _ = empty.PageSetup.Footer; _ = empty.PageSetup.FirstPageFooter; _ = empty.PageSetup.EvenPageFooter;
        var afterEmpty = source.AddSection();
        afterEmpty._sectionProperties.RemoveAllChildren<HeaderReference>();
        afterEmpty._sectionProperties.RemoveAllChildren<FooterReference>();
        using var stream = new MemoryStream();
        source.Save(stream);
        stream.Position = 0;
        using var document = WordDocument.Load(stream, new WordLoadOptions { PersistenceMode = DocumentPersistenceMode.Explicit });
        var section = document.Sections[1];
        foreach (var role in new[] { WordHeaderFooterType.Default, WordHeaderFooterType.First, WordHeaderFooterType.Even }) {
            Assert.Null(section.PageSetup.GetHeader(role));
            Assert.Null(section.PageSetup.GetFooter(role));
        }
        string xml = document._wordprocessingDocument.MainDocumentPart!.Document!.OuterXml;
        var literal = section.Find("needle", new OfficeTextSearchOptions { WholeWords = true });
        var regex = section.Find(new Regex(@"\bneedle\b"));
        foreach (var result in new[] { literal, regex }) {
            Assert.Equal(14, result.MatchCount);
            Assert.Equal("needle selected body", Assert.Single(result.BodyParagraphs).Text);
            Assert.Equal("needle selected table", Assert.Single(result.TableParagraphs).Text);
            var groups = new[] { result.DefaultHeaderParagraphs, result.FirstPageHeaderParagraphs, result.EvenPageHeaderParagraphs,
                result.DefaultFooterParagraphs, result.FirstPageFooterParagraphs, result.EvenPageFooterParagraphs };
            for (int index = 0; index < groups.Length; index++)
                Assert.Equal(new[] { "needle story " + index, "needle story table " + index }, groups[index].Select(p => p.Text));
            Assert.DoesNotContain(result.Paragraphs, p => p.Text == "needle outside");
        }
        foreach (var sectionWithEmptyStories in document.Sections.Skip(2)) {
            Assert.Equal(0, sectionWithEmptyStories.Find("needle").MatchCount);
            Assert.Equal(0, sectionWithEmptyStories.Find(new Regex("needle")).MatchCount);
        }
        Assert.Equal(15, document.Find("needle", new OfficeTextSearchOptions { WholeWords = true }).MatchCount);
        Assert.Equal(xml, document._wordprocessingDocument.MainDocumentPart!.Document!.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void TableControlsRetainTheirNativeFamiliesWithoutGenericDuplicates(bool nested) {
        string root = Path.Combine(Path.GetTempPath(), "word-table-controls-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string picture = Path.Combine(root, "pixel.png");
            File.WriteAllBytes(picture, Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP4/w8AAv8B/h10yjMAAAAASUVORK5CYII="));
            using var source = WordDocument.Create();
            var outer = source.AddTable(1, 1);
            var table = nested ? outer.Rows[0].Cells[0].AddTable(1, 1) : outer;
            var cell = table.Rows[0].Cells[0];
            cell.AddParagraph().AddStructuredDocumentTag("Text value", "Text", "Text");
            cell.AddParagraph().AddCheckBox(false, "Check", "Check");
            cell.AddParagraph().AddDatePicker(new DateTime(2026, 1, 1), "Date", "Date");
            cell.AddParagraph().AddDropDownList(new[] { "Low", "High" }, "DropDown", "DropDown");
            cell.AddParagraph().AddComboBox(new[] { "Email", "Phone" }, "Combo", "Combo");
            cell.AddParagraph().AddPictureControl(picture, OfficeLength.Pixels(1), OfficeLength.Pixels(1), "Picture", "Picture");
            cell.AddParagraph().AddRepeatingSection("Item", "Repeating", "Repeating");
            using var stream = new MemoryStream();
            source.Save(stream);
            stream.Position = 0;
            using var document = WordDocument.Load(stream, new WordLoadOptions { PersistenceMode = DocumentPersistenceMode.Explicit });
            string xml = document._wordprocessingDocument.MainDocumentPart!.Document!.OuterXml;
            Assert.Equal("Text", Assert.Single(document.TextContentControls).Tag);
            Assert.Equal("Check", Assert.Single(document.CheckBoxes).Tag);
            Assert.Equal("Date", Assert.Single(document.DatePickers).Tag);
            Assert.Equal("DropDown", Assert.Single(document.DropDownLists).Tag);
            Assert.Equal("Combo", Assert.Single(document.ComboBoxes).Tag);
            Assert.Equal("Picture", Assert.Single(document.PictureControls).Tag);
            Assert.Equal("Repeating", Assert.Single(document.RepeatingSections).Tag);
            var section = document.Sections[0];
            Assert.Single(section.CheckBoxes); Assert.Single(section.DatePickers); Assert.Single(section.DropDownLists);
            Assert.Single(section.ComboBoxes); Assert.Single(section.PictureControls); Assert.Single(section.RepeatingSections);
            var rootTable = Assert.Single(section.Tables);
            Assert.Equal(7, rootTable.StructuredDocumentTags.Count);
            Assert.Single(rootTable.CheckBoxes); Assert.Single(rootTable.DatePickers);
            Assert.Single(rootTable.DropDownLists); Assert.Single(rootTable.RepeatingSections);
            var specialized = document.CheckBoxes.Select(c => c._sdtRun).Concat(document.DatePickers.Select(c => c._sdtRun))
                .Concat(document.DropDownLists.Select(c => c._sdtRun)).Concat(document.ComboBoxes.Select(c => c._sdtRun))
                .Concat(document.PictureControls.Select(c => c._sdtRun)).Concat(document.RepeatingSections.Select(c => c._sdtRun)).ToArray();
            Assert.Equal(6, specialized.Distinct().Count());
            Assert.DoesNotContain(document.TextContentControls, c => specialized.Contains(c.SdtElement));
            Assert.Equal(xml, document._wordprocessingDocument.MainDocumentPart!.Document!.OuterXml);
        } finally { Directory.Delete(root, true); }
    }
}
