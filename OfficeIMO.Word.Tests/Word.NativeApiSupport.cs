using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class NativeWordApiSupportTests {
    [Fact]
    public void PathRedlinesRetainDefaultAndExplicitNativeSettings() {
        string root = Path.Combine(Path.GetTempPath(), "word-redline-paths-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string original = Path.Combine(root, "original.docx"), modified = Path.Combine(root, "modified.docx");
            using (var source = WordDocument.Create()) { source.AddParagraph("Before"); source.Save(original); }
            using (var target = WordDocument.Create()) { target.AddParagraph("After"); target.Save(modified); }
            string defaults = Path.Combine(root, "defaults.docx"), explicitOptions = Path.Combine(root, "explicit.docx");
            WordDocumentComparer.CreateRedlineDocument(original, modified, defaults);
            WordDocumentComparer.CreateRedlineDocument(original, modified, explicitOptions, new WordComparisonRedlineOptions { Author = "Path contract" });
            foreach (var item in new[] { (defaults, "OfficeIMO"), (explicitOptions, "Path contract") }) {
                using var redline = WordDocument.Load(item.Item1);
                var revisions = redline._wordprocessingDocument.MainDocumentPart!.Document!.Descendants<InsertedRun>().ToArray();
                Assert.NotEmpty(revisions);
                Assert.All(revisions, revision => Assert.Equal(item.Item2, revision.Author!.Value));
                Assert.Empty(redline.ValidateDocument());
            }
        } finally { Directory.Delete(root, true); }
    }

    [Fact]
    public void SectionFindKeepsNativePartGroupsAndExcludesOtherSections() {
        using var document = WordDocument.Create();
        document.AddParagraph("needle outside the selected section");
        var section = document.AddSection();
        section.AddParagraph("needle body needles");
        section.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0].Text = "NEEDLE table";
        section.PageSetup.Header.AddParagraph("needle header");
        section.PageSetup.Footer.AddParagraph("needle footer");
        var literal = section.Find("needle", new OfficeTextSearchOptions { WholeWords = true });
        Assert.Equal(4, literal.MatchCount);
        Assert.Equal("needle body needles", Assert.Single(literal.BodyParagraphs).Text);
        Assert.Equal("NEEDLE table", Assert.Single(literal.TableParagraphs).Text);
        Assert.Equal("needle header", Assert.Single(literal.DefaultHeaderParagraphs).Text);
        Assert.Equal("needle footer", Assert.Single(literal.DefaultFooterParagraphs).Text);
        Assert.DoesNotContain(literal.Paragraphs, paragraph => paragraph.Text.Contains("outside"));
        var regex = section.Find(new Regex(@"\bneedle\b", RegexOptions.IgnoreCase));
        Assert.Equal(4, regex.MatchCount);
        Assert.Equal(literal.Paragraphs.Select(paragraph => paragraph.Text), regex.Paragraphs.Select(paragraph => paragraph.Text));
        var exact = section.Find("needle", new OfficeTextSearchOptions { MatchCase = true, WholeWords = true });
        Assert.Equal(3, exact.MatchCount);
        Assert.Empty(exact.TableParagraphs);
        Assert.Equal("needle outside the selected section", document.Sections[0].Paragraphs[0].Text);
    }

    [Fact]
    public void TextContentControlsExcludeTheSixSpecializedNativeControlFamilies() {
        string root = Path.Combine(Path.GetTempPath(), "word-controls-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string picture = Path.Combine(root, "pixel.png");
            File.WriteAllBytes(picture, Convert.FromBase64String("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP4/w8AAv8B/h10yjMAAAAASUVORK5CYII="));
            using var document = WordDocument.Create();
            document.AddStructuredDocumentTag("Text value", "Text", "Text");
            document.AddParagraph().AddCheckBox(false, "Check", "Check");
            var date = document.AddParagraph().AddDatePicker(new DateTime(2026, 1, 1), "Date", "Date");
            Assert.Null(date.Text);
            date.Date = new DateTime(2026, 1, 1);
            document.AddParagraph().AddDropDownList(new[] { "Low", "High" }, "DropDown", "DropDown").SelectedValue = "High";
            document.AddParagraph().AddComboBox(new[] { "Email", "Phone" }, "Combo", "Combo", defaultValue: "Email");
            document.AddParagraph().AddPictureControl(picture, OfficeLength.Pixels(1), OfficeLength.Pixels(1), "Picture", "Picture");
            document.AddParagraph().AddRepeatingSection("Item", "Repeating", "Repeating");
            Assert.Equal("Text", Assert.Single(document.TextContentControls).Tag);
            var controls = document.TextContentControls.Cast<object>().Concat(document.CheckBoxes).Concat(document.DatePickers)
                .Concat(document.DropDownLists).Concat(document.ComboBoxes).Concat(document.PictureControls).Concat(document.RepeatingSections).ToArray();
            Assert.Equal(7, controls.Length);
            Assert.Contains(controls, control => control is WordStructuredDocumentTag);
            Assert.Contains(controls, control => control is WordCheckBox);
            Assert.Contains(controls, control => control is WordDatePicker);
            Assert.Contains(controls, control => control is WordDropDownList);
            Assert.Contains(controls, control => control is WordComboBox);
            Assert.Contains(controls, control => control is WordPictureControl);
            Assert.Contains(controls, control => control is WordRepeatingSection);
            Assert.Equal("☐", Assert.Single(document.CheckBoxes).Text);
            Assert.Equal("2026-01-01", Assert.Single(document.DatePickers).Text);
            Assert.Equal("High", Assert.Single(document.DropDownLists).Text);
            Assert.Equal("Email", Assert.Single(document.ComboBoxes).Text);
            Assert.Null(Assert.Single(document.PictureControls).Text);
            Assert.Null(Assert.Single(document.RepeatingSections).Text);
        } finally { Directory.Delete(root, true); }
    }

    [Theory]
    [InlineData(WordComparisonRedlineMode.ReportArtifact)]
    [InlineData(WordComparisonRedlineMode.InPlaceTarget)]
    public void ModelRedlinesUseUnsavedEncryptedInputsAndPreserveCallerState(WordComparisonRedlineMode mode) {
        string root = Path.Combine(Path.GetTempPath(), "word-redline-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(root);
        try {
            string sourcePath = Path.Combine(root, "source.docx"), targetPath = Path.Combine(root, "target.docx"), output = Path.Combine(root, "redline.docx");
            using (var created = WordDocument.Create()) {
                created.AddParagraph("source STORED"); created.AddParagraph("UNCHANGED CASE");
                created.Save(sourcePath, new WordSaveOptions { Password = "source-software-fixture" });
            }
            using (var created = WordDocument.Create()) {
                created.AddParagraph("target STORED"); created.AddParagraph("unchanged case");
                created.Save(targetPath, new WordSaveOptions { Password = "target-software-fixture" });
            }
            byte[] sourceBytes = File.ReadAllBytes(sourcePath), targetBytes = File.ReadAllBytes(targetPath);
            using var source = WordDocument.Load(sourcePath, new WordLoadOptions { Password = "source-software-fixture", PersistenceMode = DocumentPersistenceMode.Explicit });
            using var target = WordDocument.Load(targetPath, new WordLoadOptions { Password = "target-software-fixture", PersistenceMode = DocumentPersistenceMode.Explicit });
            source.Paragraphs[0].Text = "source LIVE";
            target.Paragraphs[0].Text = "target LIVE";
            string sourceXml = source._wordprocessingDocument.MainDocumentPart!.Document!.OuterXml;
            string targetXml = target._wordprocessingDocument.MainDocumentPart!.Document!.OuterXml;
            var comparison = new WordComparisonOptions { IgnoreCase = true, IncludedScopes = new HashSet<WordComparisonScope> { WordComparisonScope.Paragraph, WordComparisonScope.Run } };
            var options = new WordComparisonRedlineOptions { Mode = mode, Author = "Native contract", DateTime = new DateTime(2026, 11, 4, 12, 0, 0, DateTimeKind.Utc), ComparisonOptions = comparison };
            var result = WordDocumentComparer.CreateRedlineDocument(source, target, output, options);
            Assert.Contains(result.Findings, finding => (finding.SourceText ?? string.Empty).Contains("source LIVE"));
            Assert.DoesNotContain(result.Findings, finding => finding.SourceText == "UNCHANGED CASE" || finding.TargetText == "unchanged case");
            using (var redline = WordDocument.Load(output, new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly })) {
                var body = redline._wordprocessingDocument.MainDocumentPart!.Document!.Body!;
                Assert.Contains("source LIVE", string.Concat(body.Descendants<DeletedRun>().Select(run => run.InnerText)));
                Assert.Contains("target LIVE", string.Concat(body.Descendants<InsertedRun>().Select(run => run.InnerText)));
                Assert.All(body.Descendants<InsertedRun>(), run => Assert.Equal("Native contract", run.Author!.Value));
                Assert.All(body.Descendants<DeletedRun>(), run => Assert.Equal("Native contract", run.Author!.Value));
                Assert.All(body.Descendants<InsertedRun>(), run => Assert.Equal(options.DateTime, run.Date?.Value));
                Assert.All(body.Descendants<DeletedRun>(), run => Assert.Equal(options.DateTime, run.Date?.Value));
                Assert.Empty(redline.ValidateDocument());
            }
            Assert.Equal(sourceXml, source._wordprocessingDocument.MainDocumentPart!.Document!.OuterXml);
            Assert.Equal(targetXml, target._wordprocessingDocument.MainDocumentPart!.Document!.OuterXml);
            Assert.Equal(sourceBytes, File.ReadAllBytes(sourcePath));
            Assert.Equal(targetBytes, File.ReadAllBytes(targetPath));
            Assert.Same(comparison, options.ComparisonOptions);
            Assert.True(comparison.IgnoreCase);
            Assert.Equal(mode, options.Mode);
            Assert.Equal("Native contract", options.Author);
            Assert.Throws<InvalidOperationException>(() => WordDocumentComparer.CreateRedlineDocument(source, target, sourcePath, options));
            Assert.Throws<InvalidOperationException>(() => WordDocumentComparer.CreateRedlineDocument(source, target, targetPath, options));
            source.AddParagraph("Still owned by the caller");
            target.AddParagraph("Still owned by the caller");
            Assert.Equal(new[] { "redline.docx", "source.docx", "target.docx" }, Directory.GetFiles(root).Select(Path.GetFileName).OrderBy(name => name));
        } finally { Directory.Delete(root, true); }
    }
}
