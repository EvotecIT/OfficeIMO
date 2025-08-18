using System.IO;
using System.Linq;
using System.Reflection;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    /// <summary>
    /// Contains tests for custom paragraph styles.
    /// </summary>
    public partial class Word {
        [Fact]
        public void Test_RegisterCustomParagraphStyle() {
            var style = new Style { Type = StyleValues.Paragraph, StyleId = "MyStyle" };
            WordParagraphStyle.RegisterCustomStyle("MyStyle", style);

            string filePath = Path.Combine(_directoryWithFiles, "CustomParagraphStyle.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Text").SetStyleId("MyStyle");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var styles = document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles;
                Assert.NotNull(styles.Elements<Style>().FirstOrDefault(s => s.StyleId == "MyStyle"));
            }

            var field = typeof(WordParagraphStyle).GetField("_customStyles", BindingFlags.NonPublic | BindingFlags.Static)!;
            var dict = (IDictionary<string, Style>)field.GetValue(null)!;
            dict.Clear();
        }

        [Fact]
        public void Test_OverrideBuiltInParagraphStyle() {
            var original = WordParagraphStyle.GetStyleDefinition(WordParagraphStyles.Normal);
            var custom = new Style { Type = StyleValues.Paragraph, StyleId = "Normal" };
            WordParagraphStyle.OverrideBuiltInStyle(WordParagraphStyles.Normal, custom);

            Assert.Equal(custom, WordParagraphStyle.GetStyleDefinition(WordParagraphStyles.Normal));

            WordParagraphStyle.OverrideBuiltInStyle(WordParagraphStyles.Normal, original);
        }

        [Fact]
        public void Test_LoadDocumentWithExistingCustomStyle() {
            var style = new Style { Type = StyleValues.Paragraph, StyleId = "MyStyle" };
            style.Append(new StyleName { Val = "Original" });
            WordParagraphStyle.RegisterCustomStyle("MyStyle", style);

            string filePath = Path.Combine(_directoryWithFiles, "CustomStylePreserve.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Text").SetStyleId("MyStyle");
                document.Save();
            }

            var updated = new Style { Type = StyleValues.Paragraph, StyleId = "MyStyle" };
            updated.Append(new StyleName { Val = "Updated" });
            WordParagraphStyle.RegisterCustomStyle("MyStyle", updated);

            using (WordDocument document = WordDocument.Load(filePath)) {
                var styles = document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles;
                var loaded = styles.Elements<Style>().First(s => s.StyleId == "MyStyle");
                Assert.Equal("Original", loaded.StyleName.Val);
            }

            using (WordDocument document = WordDocument.Load(filePath, overrideStyles: true)) {
                var styles = document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles;
                var loaded = styles.Elements<Style>().First(s => s.StyleId == "MyStyle");
                Assert.Equal("Updated", loaded.StyleName.Val);
            }

            // cleanup
            var field = typeof(WordParagraphStyle).GetField("_customStyles", BindingFlags.NonPublic | BindingFlags.Static)!;
            var dict = (IDictionary<string, Style>)field.GetValue(null)!;
            dict.Clear();
        }

        [Fact]
        public void Test_LoadDocumentWithExistingCustomStyle_ReadOnly() {
            var style = new Style { Type = StyleValues.Paragraph, StyleId = "MyStyle" };
            style.Append(new StyleName { Val = "Original" });
            WordParagraphStyle.RegisterCustomStyle("MyStyle", style);

            string filePath = Path.Combine(_directoryWithFiles, "CustomStyleReadOnly.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Text").SetStyleId("MyStyle");
                document.Save();
            }

            var updated = new Style { Type = StyleValues.Paragraph, StyleId = "MyStyle" };
            updated.Append(new StyleName { Val = "Updated" });
            WordParagraphStyle.RegisterCustomStyle("MyStyle", updated);

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true, overrideStyles: true)) {
                var styles = document._wordprocessingDocument.MainDocumentPart.StyleDefinitionsPart.Styles;
                var loaded = styles.Elements<Style>().First(s => s.StyleId == "MyStyle");
                Assert.Equal("Original", loaded.StyleName.Val);
            }

            // cleanup
            var field = typeof(WordParagraphStyle).GetField("_customStyles", BindingFlags.NonPublic | BindingFlags.Static)!;
            var dict = (IDictionary<string, Style>)field.GetValue(null)!;
            dict.Clear();
        }
    }
}
