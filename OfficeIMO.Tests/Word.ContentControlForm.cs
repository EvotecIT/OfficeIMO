using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_ContentControlFormValuesFillAndExtractByTagOrAlias() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithContentControlFormValues.docx");
            string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");
            string replacementImagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddStructuredDocumentTag("Contoso", "Client Alias", "ClientName");
                document.AddStructuredDocumentTag("Initial notes", "Notes Alias");
                document.AddParagraph("Accepted:").AddCheckBox(false, "Accepted Alias", "Accepted");
                document.AddParagraph("Due:").AddDatePicker(new DateTime(2026, 1, 1), "Due Alias", "DueDate");
                document.AddParagraph("Priority:").AddDropDownList(new[] { "Low", "Medium", "High" }, "Priority Alias", "Priority");
                document.AddParagraph("Contact:").AddComboBox(new[] { "Email", "Phone" }, "Contact Alias", "ContactMethod", defaultValue: "Email");
                document.AddParagraph("Logo:").AddPictureControl(imagePath, 24, 24, "Logo Alias", "Logo");
                document.AddParagraph("Items:").AddRepeatingSection("Items", "Items Alias", "Items");

                int updated = document.FillContentControlValues(new Dictionary<string, object?> {
                    ["ClientName"] = "Northwind",
                    ["Notes Alias"] = "Alias fallback notes",
                    ["Accepted"] = "yes",
                    ["DueDate"] = "2026-05-29",
                    ["Priority"] = "High",
                    ["ContactMethod"] = "Phone",
                    ["Logo"] = WordContentControlPictureValue.FromFile(replacementImagePath),
                    ["Items"] = new[] { "First", "Second", "Third" }
                });

                Assert.Equal(8, updated);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Dictionary<string, object?> values = document.ExtractContentControlValues();

                Assert.Equal("Northwind", values["ClientName"]);
                Assert.Equal("Alias fallback notes", values["Notes Alias"]);
                Assert.Equal(true, values["Accepted"]);
                Assert.Equal(new DateTime(2026, 5, 29), ((DateTime?)values["DueDate"])!.Value.Date);
                Assert.Equal("High", values["Priority"]);
                Assert.Equal("Phone", values["ContactMethod"]);
                var logo = Assert.IsType<WordContentControlPictureValue>(values["Logo"]);
                Assert.Equal("EvotecLogo.png", logo.FileName);
                Assert.Equal(File.ReadAllBytes(replacementImagePath), logo.Bytes);
                var items = Assert.IsAssignableFrom<IReadOnlyList<string>>(values["Items"]);
                Assert.Equal(new[] { "First", "Second", "Third" }, items);

                Assert.Equal("Northwind", document.GetStructuredDocumentTagByTag("ClientName")!.Text);
                Assert.True(document.GetCheckBoxByTag("Accepted")!.IsChecked);
                Assert.Equal(new DateTime(2026, 5, 29), document.GetDatePickerByTag("DueDate")!.Date!.Value.Date);
                Assert.Equal("High", document.GetDropDownListByTag("Priority")!.SelectedValue);
                Assert.Equal("Phone", document.GetComboBoxByTag("ContactMethod")!.SelectedValue);
                Assert.NotNull(document.GetPictureControlByTag("Logo")!.Image);
                Assert.Equal(new[] { "First", "Second", "Third" }, document.GetRepeatingSectionByTag("Items")!.TextItems);
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false)) {
                var textValues = wordDocument.MainDocumentPart!.Document.Body!.Descendants<SdtRun>()
                    .SelectMany(sdt => sdt.SdtContentRun?.Descendants<Text>() ?? Enumerable.Empty<Text>())
                    .Select(text => text.Text)
                    .ToArray();

                Assert.Contains("2026-05-29", textValues);
                Assert.Contains("High", textValues);
                Assert.Contains("Phone", textValues);

                var repeatingItems = wordDocument.MainDocumentPart!.Document.Body!.Descendants()
                    .Where(element => element.LocalName == "repeatingSectionItem")
                    .ToList();
                Assert.True(repeatingItems.Count >= 3);
                Assert.Contains(repeatingItems, item => item.InnerText.Contains("First"));
                Assert.Contains(repeatingItems, item => item.InnerText.Contains("Second"));
                Assert.Contains(repeatingItems, item => item.InnerText.Contains("Third"));
            }
        }

        [Fact]
        public void Test_ContentControlFormValidationReportsMissingInvalidAndUnusedValues() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithContentControlFormValidation.docx");
            string imagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddStructuredDocumentTag("Contoso", "Client Alias", "ClientName");
                document.AddParagraph("Accepted:").AddCheckBox(false, "Accepted Alias", "Accepted");
                document.AddParagraph("Due:").AddDatePicker(new DateTime(2026, 1, 1), "Due Alias", "DueDate");
                document.AddParagraph("Priority:").AddDropDownList(new[] { "Low", "Medium", "High" }, "Priority Alias", "Priority");
                document.AddParagraph("Contact:").AddComboBox(new[] { "Email", "Phone" }, "Contact Alias", "ContactMethod", defaultValue: "Email");
                document.AddParagraph("Logo:").AddPictureControl(imagePath, 24, 24, "Logo Alias", "Logo");
                document.AddParagraph("Items:").AddRepeatingSection("Items", "Items Alias", "Items");

                WordContentControlFormValidationResult valid = document.ValidateContentControlValues(new Dictionary<string, object?> {
                    ["clientname"] = "Northwind",
                    ["Accepted"] = true,
                    ["DueDate"] = new DateTime(2026, 5, 29),
                    ["Priority"] = "Medium",
                    ["ContactMethod"] = "Phone",
                    ["Logo"] = WordContentControlPictureValue.FromFile(imagePath),
                    ["Items"] = new[] { "One", "Two" }
                });

                Assert.True(valid.IsValid, string.Join(Environment.NewLine, valid.Issues.Select(issue => issue.Message)));
                Assert.Equal(new[] { "Accepted", "ClientName", "ContactMethod", "DueDate", "Items", "Logo", "Priority" }, valid.ExpectedKeys);
                Assert.Same(valid, valid.EnsureValid());

                WordContentControlFormValidationResult invalid = document.ValidateContentControlValues(new Dictionary<string, object?> {
                    ["ClientName"] = "Northwind",
                    ["Accepted"] = "not-a-bool",
                    ["DueDate"] = "not-a-date",
                    ["Priority"] = "Urgent",
                    ["Logo"] = WordContentControlPictureValue.FromFile(Path.Combine(_directoryWithImages, "missing-logo.png")),
                    ["Items"] = 12345,
                    ["Extra"] = "Ignored"
                });

                Assert.False(invalid.IsValid);
                Assert.Contains(invalid.Issues, issue => issue.Kind == WordContentControlFormIssueKind.InvalidBoolean && issue.Key == "Accepted");
                Assert.Contains(invalid.Issues, issue => issue.Kind == WordContentControlFormIssueKind.InvalidDate && issue.Key == "DueDate");
                Assert.Contains(invalid.Issues, issue => issue.Kind == WordContentControlFormIssueKind.InvalidChoice && issue.Key == "Priority");
                Assert.Contains(invalid.Issues, issue => issue.Kind == WordContentControlFormIssueKind.InvalidImage && issue.Key == "Logo");
                Assert.Contains(invalid.Issues, issue => issue.Kind == WordContentControlFormIssueKind.InvalidRepeatingSection && issue.Key == "Items");
                Assert.Contains(invalid.Issues, issue => issue.Kind == WordContentControlFormIssueKind.MissingValue && issue.Key == "ContactMethod");
                Assert.Contains(invalid.Issues, issue => issue.Kind == WordContentControlFormIssueKind.UnusedValue && issue.Key == "Extra");
                Assert.Throws<InvalidOperationException>(() => invalid.EnsureValid());
            }
        }

        [Fact]
        public void Test_ContentControlFormValidationReportsAmbiguousTagAliasKeys() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithAmbiguousContentControlFormKeys.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddStructuredDocumentTag("Contoso", "Shared Alias", "ClientName");
                document.AddStructuredDocumentTag("Migration", "Shared Alias", "ProjectName");

                WordContentControlFormValidationResult result = document.ValidateContentControlValues(new Dictionary<string, object?> {
                    ["ClientName"] = "Northwind",
                    ["ProjectName"] = "Website refresh"
                });

                WordContentControlFormIssue duplicate = Assert.Single(result.Issues, issue => issue.Kind == WordContentControlFormIssueKind.DuplicateKey);
                Assert.Equal("Shared Alias", duplicate.Key);
                Assert.Contains("multiple content controls", duplicate.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Throws<InvalidOperationException>(() => result.EnsureValid());
            }
        }
    }
}
