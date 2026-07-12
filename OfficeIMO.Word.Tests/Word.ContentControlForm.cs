using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
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
                document.Save();
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
                Assert.DoesNotContain("Accepted Alias", values.Keys);
                Assert.DoesNotContain("Due Alias", values.Keys);
                Assert.DoesNotContain("Priority Alias", values.Keys);
                Assert.DoesNotContain("Contact Alias", values.Keys);
                Assert.DoesNotContain("Logo Alias", values.Keys);
                Assert.DoesNotContain("Items Alias", values.Keys);

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
        public void Test_ContentControlForm_WordAuthoredFixtureCanValidateFillAndExtractValues() {
            string filePath = CopyFixtureDoc(Path.Combine("Word", "PremiumGaps", "TemplateMailMerge", "word-authored-content-control-form.docx"));
            string replacementImagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordContentControlFormValidationResult validation = document.ValidateContentControlValues(new Dictionary<string, object?> {
                    ["ClientName"] = "Northwind Traders",
                    ["Accepted"] = true,
                    ["DueDate"] = new DateTime(2026, 7, 15),
                    ["Priority"] = "High",
                    ["ContactMethod"] = "Teams",
                    ["Notes"] = "Imported rich-text notes",
                    ["Logo"] = WordContentControlPictureValue.FromFile(replacementImagePath),
                    ["ReferenceCode"] = "REF-2026-001"
                });

                Assert.True(validation.IsValid, string.Join(Environment.NewLine, validation.Issues.Select(issue => issue.Message)));
                Assert.Equal(new[] { "Accepted", "ClientName", "ContactMethod", "DueDate", "Logo", "Notes", "Priority", "ReferenceCode" }, validation.ExpectedKeys);
                Assert.Contains("\"expectedKeyCount\": 8", validation.ToJson(), StringComparison.Ordinal);
                Assert.Contains("- ReferenceCode", validation.ToMarkdown(), StringComparison.Ordinal);

                int updated = document.FillContentControlValues(new Dictionary<string, object?> {
                    ["ClientName"] = "Northwind Traders",
                    ["Accepted"] = true,
                    ["DueDate"] = new DateTime(2026, 7, 15),
                    ["Priority"] = "High",
                    ["ContactMethod"] = "Teams",
                    ["Notes"] = "Imported rich-text notes",
                    ["Logo"] = WordContentControlPictureValue.FromFile(replacementImagePath),
                    ["ReferenceCode"] = "REF-2026-001"
                });

                Assert.Equal(8, updated);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Dictionary<string, object?> values = document.ExtractContentControlValues();

                Assert.Equal("Northwind Traders", values["ClientName"]);
                Assert.Equal(true, values["Accepted"]);
                Assert.Equal(new DateTime(2026, 7, 15), ((DateTime?)values["DueDate"])!.Value.Date);
                Assert.Equal("High", values["Priority"]);
                Assert.Equal("Teams", values["ContactMethod"]);
                Assert.Equal("Imported rich-text notes", values["Notes"]);
                var logo = Assert.IsType<WordContentControlPictureValue>(values["Logo"]);
                Assert.Equal("EvotecLogo.png", logo.FileName);
                Assert.Equal(File.ReadAllBytes(replacementImagePath), logo.Bytes);
                Assert.Equal("REF-2026-001", values["ReferenceCode"]);
                Assert.True(document.GetCheckBoxByTag("Accepted")!.IsChecked);
                Assert.Equal("High", document.GetDropDownListByTag("Priority")!.SelectedValue);
                Assert.Equal("Teams", document.GetComboBoxByTag("ContactMethod")!.SelectedValue);
                Assert.Equal("Imported rich-text notes", document.GetStructuredDocumentTagByTag("Notes")!.Text);
                Assert.Equal(File.ReadAllBytes(replacementImagePath), document.GetPictureControlByTag("Logo")!.Image!.GetBytes());
                Assert.Equal("REF-2026-001", document.GetStructuredDocumentTagByTag("ReferenceCode")!.Text);
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false)) {
                var validator = new OpenXmlValidator(FileFormatVersions.Microsoft365);
                var errors = validator.Validate(wordDocument).ToList();
                Assert.True(errors.Count == 0, FormatValidationErrors(errors));
            }
        }

        [Fact]
        public void Test_ContentControlFormValidationAcceptsListItemValues() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithContentControlFormListItemValues.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordDropDownList country = document.AddParagraph("Country:").AddDropDownList(new[] { "United States" }, "Country Alias", "Country");
                ListItem countryItem = country._sdtRun.SdtProperties!.GetFirstChild<SdtContentDropDownList>()!.Elements<ListItem>().Single();
                countryItem.Value = "US";
                countryItem.DisplayText = "United States";

                WordComboBox status = document.AddParagraph("Status:").AddComboBox(new[] { "In progress" }, "Status Alias", "Status", defaultValue: "In progress");
                ListItem statusItem = status._sdtRun.SdtProperties!.GetFirstChild<SdtContentComboBox>()!.Elements<ListItem>().Single();
                statusItem.Value = "IP";
                statusItem.DisplayText = "In progress";

                WordContentControlFormValidationResult validation = document.ValidateContentControlValues(new Dictionary<string, object?> {
                    ["Country"] = "US",
                    ["Status"] = "IP"
                });

                Assert.True(validation.IsValid, validation.ToJson());
                int updated = document.FillContentControlValues(new Dictionary<string, object?> {
                    ["Country"] = "US",
                    ["Status"] = "IP"
                });

                Assert.Equal(2, updated);
                Assert.Equal("US", country.SelectedValue);
                Assert.Equal("IP", status.SelectedValue);
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

                string json = invalid.ToJson();
                using JsonDocument parsed = JsonDocument.Parse(json);
                JsonElement root = parsed.RootElement;
                Assert.False(root.GetProperty("isValid").GetBoolean());
                Assert.Equal(7, root.GetProperty("expectedKeyCount").GetInt32());
                Assert.Equal(7, root.GetProperty("suppliedKeyCount").GetInt32());
                Assert.Equal(invalid.Issues.Count, root.GetProperty("issueCount").GetInt32());
                Assert.Contains(root.GetProperty("expectedKeys").EnumerateArray(), key => key.GetString() == "ContactMethod");
                Assert.Contains(root.GetProperty("issues").EnumerateArray(), issue =>
                    issue.GetProperty("kind").GetString() == nameof(WordContentControlFormIssueKind.MissingValue) &&
                    issue.GetProperty("key").GetString() == "ContactMethod" &&
                    issue.GetProperty("controlType").GetString() == "Combo box");

                string markdown = invalid.ToMarkdown();
                Assert.Contains("# Content-Control Form Validation", markdown);
                Assert.Contains("| Valid | no |", markdown);
                Assert.Contains("## Expected Keys", markdown);
                Assert.Contains("- ContactMethod", markdown);
                Assert.Contains("| MissingValue | ContactMethod | Combo box |", markdown);
            }
        }

        [Fact]
        public void Test_ContentControlFormPictureValuesDoNotReadRawStringPaths() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithRawPathPictureFormValue.docx");
            string originalImagePath = Path.Combine(_directoryWithImages, "Kulek.jpg");
            string replacementImagePath = Path.Combine(_directoryWithImages, "EvotecLogo.png");
            byte[] originalBytes;

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Logo:").AddPictureControl(originalImagePath, 24, 24, "Logo Alias", "Logo");
                originalBytes = document.GetPictureControlByTag("Logo")!.Image!.GetBytes();

                WordContentControlFormValidationResult validation = document.ValidateContentControlValues(new Dictionary<string, object?> {
                    ["Logo"] = replacementImagePath
                });

                Assert.False(validation.IsValid);
                Assert.Contains(validation.Issues, issue => issue.Kind == WordContentControlFormIssueKind.InvalidImage && issue.Key == "Logo");

                int updated = document.FillContentControlValues(new Dictionary<string, object?> {
                    ["Logo"] = replacementImagePath
                });

                Assert.Equal(0, updated);
                Assert.Equal(originalBytes, document.GetPictureControlByTag("Logo")!.Image!.GetBytes());

                int trustedUpdated = document.FillContentControlValues(new Dictionary<string, object?> {
                    ["Logo"] = WordContentControlPictureValue.FromFile(replacementImagePath)
                });

                Assert.Equal(1, trustedUpdated);
                Assert.Equal(File.ReadAllBytes(replacementImagePath), document.GetPictureControlByTag("Logo")!.Image!.GetBytes());
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

        [Fact]
        public void Test_ContentControlFormPlainControlsRemainBindableWhenKeysOverlapSpecializedControls() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithOverlappingSpecializedAndPlainContentControlKeys.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Accepted:").AddCheckBox(false, "Shared", "Accepted");
                WordStructuredDocumentTag plain = document.AddParagraph("Plain:").AddStructuredDocumentTag("Original", "Shared", "PlainShared");

                WordContentControlFormValidationResult validation = document.ValidateContentControlValues(
                    new Dictionary<string, object?> {
                        ["Shared"] = true
                    },
                    WordContentControlFormKey.Alias);

                Assert.Contains(validation.Issues, issue =>
                    issue.Kind == WordContentControlFormIssueKind.DuplicateKey &&
                    issue.Key == "Shared");

                int updated = document.FillContentControlValues(
                    new Dictionary<string, object?> {
                        ["Shared"] = "Plain value"
                    },
                    WordContentControlFormKey.Alias);

                Assert.Equal(1, updated);
                Assert.False(document.GetCheckBoxByTag("Accepted")!.IsChecked);
                Assert.Equal("Plain value", plain.Text);
            }
        }

        [Fact]
        public void Test_ContentControlForm_TableScopedSpecializedControlsRemainBindableAsGenericTags() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithTableScopedSpecializedContentControl.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                var table = new Table(
                    new TableRow(
                        new TableCell(
                            new Paragraph(
                                new SdtRun(
                                    new SdtProperties(
                                        new SdtAlias { Val = "Table Accepted" },
                                        new Tag { Val = "TableAccepted" },
                                        new DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox()),
                                    new SdtContentRun(
                                        new Run(new Text("Unchecked") { Space = SpaceProcessingModeValues.Preserve })))))));
                document._document.Body!.Append(table);

                WordContentControlFormValidationResult validation = document.ValidateContentControlValues(new Dictionary<string, object?> {
                    ["TableAccepted"] = "Approved"
                });
                Assert.True(validation.IsValid);

                int updated = document.FillContentControlValues(new Dictionary<string, object?> {
                    ["TableAccepted"] = "Approved"
                });

                Assert.Equal(1, updated);
                Dictionary<string, object?> values = document.ExtractContentControlValues();
                Assert.Equal("Approved", Assert.IsType<string>(values["TableAccepted"]));
                Assert.Contains(document._document.Body!.Descendants<Text>(), text => text.Text == "Approved");
            }
        }

        [Fact]
        public void Test_ContentControlFormValidationReportsUnmappedPlainControls() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithUnmappedPlainContentControl.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document._document.Body!.Append(new SdtBlock(
                    new SdtProperties(),
                    new SdtContentBlock(
                        new Paragraph(
                            new Run(
                                new Text("Unmapped plain control") { Space = SpaceProcessingModeValues.Preserve })))));

                WordContentControlFormValidationResult validation = document.ValidateContentControlValues(
                    new Dictionary<string, object?>(),
                    requireAllControls: true);

                WordContentControlFormIssue issue = Assert.Single(validation.Issues, item => item.Kind == WordContentControlFormIssueKind.UnmappedControl);
                Assert.Equal("Structured document tag", issue.ControlType);
                Assert.False(validation.IsValid);
            }
        }

        [Fact]
        public void Test_ContentControlFormValidationJsonEscapesControlCharacters() {
            string filePath = Path.Combine(_directoryWithFiles, "DocumentWithControlCharacterFormKeys.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddStructuredDocumentTag("Contoso", "Client", "ClientName");

                WordContentControlFormValidationResult result = document.ValidateContentControlValues(new Dictionary<string, object?> {
                    ["Bad\u0001Key"] = "Northwind"
                });

                string json = result.ToJson();

                Assert.Contains("Bad\\u0001Key", json, StringComparison.Ordinal);
                using JsonDocument parsed = JsonDocument.Parse(json);
                Assert.Equal("Bad\u0001Key", parsed.RootElement.GetProperty("suppliedKeys")[0].GetString());
            }
        }
    }
}
