using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Theory]
        [InlineData("Text", "contentLocked")]
        [InlineData("Inline", "contentLocked")]
        [InlineData("Checkbox", "contentLocked")]
        [InlineData("Date", "contentLocked")]
        [InlineData("Dropdown", "contentLocked")]
        [InlineData("Combo", "contentLocked")]
        [InlineData("Picture", "contentLocked")]
        [InlineData("Repeating", "contentLocked")]
        [InlineData("Text", "sdtContentLocked")]
        public void Test_ContentControlFormLocksRejectFillBeforeAnyMutation(string target, string lockValue) {
            using var stream = new MemoryStream();
            using (WordDocument document = WordDocument.Create(stream)) {
                AddLockTestForm(document);
                SetFormTestLock(document, target, lockValue);
                document.Save();
            }

            stream.Position = 0;
            using (WordDocument document = WordDocument.Load(stream)) {
                string before = document._document.OuterXml;
                var values = GetLockTestFormValues();
                WordContentControlFormValidationResult validation = document.ValidateContentControlValues(values);
                WordContentControlFormIssue issue = Assert.Single(validation.Issues);
                Assert.Equal(WordContentControlFormIssueKind.LockedControl, issue.Kind);
                Assert.Equal(target, issue.Key);
                Assert.Throws<InvalidOperationException>(() => document.FillContentControlValues(values));
                Assert.Equal(before, document._document.OuterXml);
                document.Save();
            }

            stream.Position = 0;
            using var package = WordprocessingDocument.Open(stream, false);
            Assert.Contains(package.MainDocumentPart!.Document.Descendants<Text>(), text => text.Text == "ORIGINAL");
            Assert.Contains(package.MainDocumentPart.Document.Descendants<Lock>(), item => item.Val!.InnerText == lockValue);
        }

        [Theory]
        [InlineData("sdtLocked")]
        [InlineData("unlocked")]
        public void Test_ContentControlFormLocksAllowEditableContent(string lockValue) {
            using var stream = new MemoryStream();
            using (WordDocument document = WordDocument.Create(stream)) {
                AddLockTestForm(document);
                foreach (string key in GetLockTestFormValues().Keys) SetFormTestLock(document, key, lockValue);
                Assert.True(document.ValidateContentControlValues(GetLockTestFormValues()).IsValid);
                Assert.Equal(8, document.FillContentControlValues(GetLockTestFormValues()));
                document.Save();
            }

            stream.Position = 0;
            using var loaded = WordDocument.Load(stream);
            Dictionary<string, object?> values = loaded.ExtractContentControlValues();
            Assert.Equal("REPLACED", values["Text"]);
            Assert.Equal("REPLACED", values["Inline"]);
            Assert.Equal(true, values["Checkbox"]);
            Assert.Equal(new DateTime(2026, 10, 8), values["Date"]);
            Assert.Equal("High", values["Dropdown"]);
            Assert.Equal("Phone", values["Combo"]);
            Assert.Equal(File.ReadAllBytes(Path.Combine(_directoryWithImages, "EvotecLogo.png")),
                Assert.IsType<WordContentControlPictureValue>(values["Picture"]).Bytes);
            Assert.Equal(new[] { "First", "Second" }, Assert.IsAssignableFrom<IReadOnlyList<string>>(values["Repeating"]));
            Assert.Equal(8, loaded._document.Descendants<Lock>().Count());
        }

        [Fact]
        public void Test_ContentControlFormLocksUseMatchedAliasAndIgnoreUnmatchedControls() {
            using var stream = new MemoryStream();
            using var document = WordDocument.Create(stream);
            document.AddStructuredDocumentTag("ORIGINAL", "Alias", "Tag");
            SetFormTestLock(document, "Tag", "contentLocked");
            Assert.True(document.ValidateContentControlValues(new Dictionary<string, object?>(), requireAllControls: false).IsValid);
            Assert.Equal(0, document.FillContentControlValues(new Dictionary<string, object?> { ["Other"] = "Ignored" }));
            var values = new Dictionary<string, object?> { ["alias"] = "REPLACED" };
            WordContentControlFormIssue issue = Assert.Single(document.ValidateContentControlValues(values).Issues);
            Assert.Equal("Alias", issue.Key);
            Assert.Throws<InvalidOperationException>(() => document.FillContentControlValues(values, WordContentControlFormKey.Alias));
            Assert.Equal("ORIGINAL", document.GetStructuredDocumentTagByTag("Tag")!.Text);
        }

        [Theory]
        [InlineData("sdtLocked")]
        [InlineData("contentLocked")]
        public void Test_ContentControlFormLocksProtectLockedRepeatingChildrenAfterReload(string lockValue) {
            using var stream = new MemoryStream();
            using (WordDocument document = WordDocument.Create(stream)) {
                document.AddParagraph().AddRepeatingSection("Items", "Items", "Items").SetTextItems(new[] { "ORIGINAL" });
                SdtRun child = document._document.Descendants<SdtRun>().First().Descendants<SdtRun>().First();
                child.SdtProperties = new SdtProperties(new Tag { Val = "Child" },
                    new Lock { Val = new DocumentFormat.OpenXml.EnumValue<LockingValues> { InnerText = lockValue } });
                document.Save();
            }
            stream.Position = 0;
            using var loaded = WordDocument.Load(stream);
            var values = new Dictionary<string, object?> { ["Items"] = new[] { "Replacement" } };
            string before = loaded._document.OuterXml;
            Assert.Contains(loaded.ValidateContentControlValues(values).Issues,
                issue => issue.Kind == WordContentControlFormIssueKind.LockedControl && issue.Key == "Items");
            Assert.Throws<InvalidOperationException>(() => loaded.FillContentControlValues(values));
            Assert.Equal(before, loaded._document.OuterXml);
        }

        private void AddLockTestForm(WordDocument document) {
            document.AddStructuredDocumentTag("ORIGINAL", "Text", "Text");
            document.AddParagraph().AddStructuredDocumentTag("ORIGINAL", "Inline", "Inline");
            document.AddParagraph().AddCheckBox(false, "Checkbox", "Checkbox");
            document.AddParagraph().AddDatePicker(new DateTime(2026, 1, 1), "Date", "Date");
            document.AddParagraph().AddDropDownList(new[] { "Low", "High" }, "Dropdown", "Dropdown");
            document.AddParagraph().AddComboBox(new[] { "Email", "Phone" }, "Combo", "Combo", defaultValue: "Email");
            document.AddParagraph().AddPictureControl(Path.Combine(_directoryWithImages, "Kulek.jpg"), 24, 24, "Picture", "Picture");
            document.AddParagraph().AddRepeatingSection("Items", "Repeating", "Repeating");
        }

        private Dictionary<string, object?> GetLockTestFormValues() => new Dictionary<string, object?> {
            ["Text"] = "REPLACED", ["Inline"] = "REPLACED", ["Checkbox"] = true,
            ["Date"] = new DateTime(2026, 10, 8), ["Dropdown"] = "High", ["Combo"] = "Phone",
            ["Picture"] = WordContentControlPictureValue.FromFile(Path.Combine(_directoryWithImages, "EvotecLogo.png")),
            ["Repeating"] = new[] { "First", "Second" }
        };

        private static void SetFormTestLock(WordDocument document, string tag, string value) {
            SdtElement element = document._document.Descendants<SdtElement>()
                .Single(item => item.GetFirstChild<SdtProperties>()?.GetFirstChild<Tag>()?.Val?.Value == tag);
            element.GetFirstChild<SdtProperties>()!.Append(new Lock { Val = new DocumentFormat.OpenXml.EnumValue<LockingValues> { InnerText = value } });
        }
    }
}
