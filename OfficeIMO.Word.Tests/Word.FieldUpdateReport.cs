using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public class WordFieldUpdateReportTests {
        private readonly string _directoryWithFiles;

        public WordFieldUpdateReportTests() {
            _directoryWithFiles = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "TempDocuments2", Guid.NewGuid().ToString("N"));
            Word.Setup(_directoryWithFiles);
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesMetadataCustomPropertiesAndFileName() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.Metadata.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Creator = "Ada Lovelace";
                document.BuiltinDocumentProperties.Title = "Premium Gap Plan";
                document.BuiltinDocumentProperties.Subject = "Field evaluator";
                document.BuiltinDocumentProperties.Keywords = "officeimo,fields";
                document.BuiltinDocumentProperties.Description = "Structured field diagnostics";
                document.BuiltinDocumentProperties.LastModifiedBy = "Grace Hopper";
                document.BuiltinDocumentProperties.Created = new DateTime(2024, 1, 2, 3, 4, 5);
                document.BuiltinDocumentProperties.Modified = new DateTime(2024, 1, 3, 4, 5, 6);
                document.CustomDocumentProperties["ClientName"] = new WordCustomProperty("Evotec");
                document.CustomDocumentProperties["Ticket"] = new WordCustomProperty(42);
                document.CustomDocumentProperties["Reviewed"] = new WordCustomProperty(true);
                document.CustomDocumentProperties["Due"] = new WordCustomProperty(new DateTime(2024, 2, 3, 12, 0, 0));

                document.AddParagraph("Author: ").AddField(WordFieldType.Author);
                document.AddParagraph("Title: ").AddField(WordFieldType.Title);
                document.AddParagraph("Subject: ").AddField(WordFieldType.Subject);
                document.AddParagraph("Keywords: ").AddField(WordFieldType.Keywords);
                document.AddParagraph("Comments: ").AddField(WordFieldType.Comments);
                document.AddParagraph("Last saved by: ").AddField(WordFieldType.LastSavedBy);
                document.AddParagraph("Created: ").AddField(WordFieldType.CreateDate);
                document.AddParagraph("Saved: ").AddField(WordFieldType.SaveDate);
                document.AddParagraph("Client: ").AddField(WordFieldType.DocProperty, parameters: new List<string> { "\"ClientName\"" });
                document.AddParagraph("Ticket: ").AddField(WordFieldType.DocProperty, parameters: new List<string> { "\"Ticket\"" });
                document.AddParagraph("Reviewed: ").AddField(WordFieldType.DocProperty, parameters: new List<string> { "\"Reviewed\"" });
                document.AddParagraph("Due: ")._paragraph.Append(BuildSimpleField(" DOCPROPERTY Due \\@ \"yyyy-MM-dd\" ", "stale-due"));
                document.AddParagraph("File: ").AddField(WordFieldType.FileName);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(13, report.TotalCount);
                Assert.Equal(13, report.UpdatedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);

                AssertUpdated(report, WordFieldType.Author, "Ada Lovelace");
                AssertUpdated(report, WordFieldType.Title, "Premium Gap Plan");
                AssertUpdated(report, WordFieldType.Subject, "Field evaluator");
                AssertUpdated(report, WordFieldType.Keywords, "officeimo,fields");
                AssertUpdated(report, WordFieldType.Comments, "Structured field diagnostics");
                AssertUpdated(report, WordFieldType.LastSavedBy, "Grace Hopper");
                AssertUpdated(report, WordFieldType.CreateDate, "2024-01-02 03:04:05");
                AssertUpdated(report, WordFieldType.SaveDate, "2024-01-03 04:05:06");
                AssertUpdated(report, WordFieldType.FileName, Path.GetFileName(filePath));
                AssertDocPropertyUpdated(report, "ClientName", "Evotec");
                AssertDocPropertyUpdated(report, "Ticket", "42");
                AssertDocPropertyUpdated(report, "Reviewed", "True");
                AssertDocPropertyUpdated(report, "Due", "2024-02-03");

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                Assert.Contains(fields, field => field.FieldType == WordFieldType.Author && field.ResultText == "Ada Lovelace");
                Assert.Contains(fields, field => field.FieldType == WordFieldType.Title && field.ResultText == "Premium Gap Plan");
                Assert.Contains(fields, field => field.FieldType == WordFieldType.FileName && field.ResultText == Path.GetFileName(filePath));
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.DocProperty &&
                    field.InstructionText.Contains("ClientName", StringComparison.Ordinal) &&
                    field.ResultText == "Evotec");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.DocProperty &&
                    field.InstructionText.Contains("Due", StringComparison.Ordinal) &&
                    field.ResultText == "2024-02-03");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesDottedMetadataAndVariableNames() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.DottedMetadata.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.CustomDocumentProperties["Client.Name"] = new WordCustomProperty("Evotec");
                document.SetDocumentVariable("Case.Id", "INC-42");
                document.AddParagraph("Client: ")._paragraph.Append(BuildSimpleField(" DOCPROPERTY Client.Name ", "stale-client"));
                document.AddParagraph("Case: ")._paragraph.Append(BuildSimpleField(" DOCVARIABLE Case.Id ", "stale-case"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(2, report.TotalCount);
                Assert.Equal(2, report.UpdatedCount);
                Assert.Equal(0, report.ParseErrorCount);
                AssertDocPropertyUpdated(report, "Client.Name", "Evotec");
                AssertDocVariableUpdated(report, "Case.Id", "INC-42");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesDocumentVariableFields() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.DocumentVariables.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.SetDocumentVariable("ClientName", "Evotec");
                document.SetDocumentVariable("Ticket", "INC-42");
                document.SetDocumentVariable("Status", "mixed status");

                document.AddParagraph("Client: ").AddField(WordFieldType.DocVariable, parameters: new List<string> { "\"ClientName\"" });
                document.AddParagraph("Ticket: ")._paragraph.Append(BuildSimpleField(" DOCVARIABLE Ticket ", "stale-ticket"));
                document.AddParagraph("Case-insensitive: ")._paragraph.Append(BuildSimpleField(" DOCVARIABLE clientname ", "stale-client-lower"));
                document.AddParagraph("Formatted: ")._paragraph.Append(BuildSimpleField(" DOCVARIABLE Status \\* Upper ", "stale-status"));
                document.AddParagraph("Missing: ")._paragraph.Append(BuildSimpleField(" DOCVARIABLE MissingVariable ", "stale-missing"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(5, report.TotalCount);
                Assert.Equal(4, report.UpdatedCount);
                Assert.Equal(1, report.SkippedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);

                AssertDocVariableUpdated(report, "ClientName", "Evotec");
                AssertDocVariableUpdated(report, "Ticket", "INC-42");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.DocVariable &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.InstructionText.Contains("clientname", StringComparison.Ordinal) &&
                    result.ResultText == "Evotec");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.DocVariable &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.InstructionText.Contains("Status", StringComparison.Ordinal) &&
                    result.ResultText == "MIXED STATUS");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.DocVariable &&
                    result.Status == WordFieldUpdateStatus.Skipped &&
                    result.InstructionText.Contains("MissingVariable", StringComparison.Ordinal) &&
                    result.ResultText == null);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.DocVariable &&
                    field.InstructionText.Contains("ClientName", StringComparison.Ordinal) &&
                    field.ResultText == "Evotec");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.DocVariable &&
                    field.InstructionText.Contains("Ticket", StringComparison.Ordinal) &&
                    field.ResultText == "INC-42");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.DocVariable &&
                    field.InstructionText.Contains("clientname", StringComparison.Ordinal) &&
                    field.ResultText == "Evotec");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.DocVariable &&
                    field.InstructionText.Contains("Status", StringComparison.Ordinal) &&
                    field.ResultText == "MIXED STATUS");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.DocVariable &&
                    field.InstructionText.Contains("MissingVariable", StringComparison.Ordinal) &&
                    field.ResultText == "stale-missing");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_PreservesSignificantResultWhitespace() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.SignificantWhitespace.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Padded quote: ")._paragraph.Append(BuildSimpleField(" QUOTE \" padded value \" ", "stale"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                AssertUpdated(report, WordFieldType.Quote, " padded value ");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                Text text = Assert.Single(document._wordprocessingDocument.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>().Single().Descendants<Text>());
                Assert.Equal(" padded value ", text.Text);
                Assert.Equal(SpaceProcessingModeValues.Preserve, text.Space?.Value);
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_CreatesResultRunForEmptySimpleFields() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.EmptySimpleField.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Creator = "Ada Lovelace";
                document.AddParagraph("Author: ")._paragraph.Append(new SimpleField { Instruction = " AUTHOR " });
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                AssertUpdated(report, WordFieldType.Author, "Ada Lovelace");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                SimpleField field = Assert.Single(document._wordprocessingDocument.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>());
                Text text = Assert.Single(field.Descendants<Text>());
                Assert.Equal("Ada Lovelace", text.Text);
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_PreservesPageCountBehaviorAndReportsUnsupportedFields() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.PagesAndUnsupported.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Page one: ").AddField(WordFieldType.Page);
                document.AddPageBreak();
                document.AddParagraph("Page two: ").AddField(WordFieldType.Page);
                document.AddParagraph("Roman page: ")._paragraph.Append(BuildSimpleField(" PAGE \\* Roman ", "stale-page-roman"));
                document.AddPageBreak();
                document.AddParagraph("Two explicit breaks")._paragraph.Append(
                    new Run(new Break { Type = BreakValues.Page }),
                    new Run(new Break { Type = BreakValues.Page }));
                document.AddParagraph("Total pages: ").AddField(WordFieldType.NumPages);
                document.AddParagraph("Formatted page: ")._paragraph.Append(BuildSimpleField(" PAGE \\# \"000\" ", "stale-page-picture"));
                document.AddParagraph("Formatted total pages: ")._paragraph.Append(BuildSimpleField(" NUMPAGES \\# \"000\" ", "stale-total-picture"));
                document.AddParagraph("Roman total pages: ")._paragraph.Append(BuildSimpleField(" NUMPAGES \\* Roman ", "stale-total-roman"));
                document.AddParagraph("Unsupported: ").AddField(WordFieldType.Database);
                document.AddParagraph()._paragraph.Append(BuildSimpleField(" SILLYFIELD value ", "Unknown"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(9, report.TotalCount);
                Assert.Equal(7, report.UpdatedCount);
                Assert.Equal(1, report.UnsupportedCount);
                Assert.Equal(1, report.ParseErrorCount);

                Assert.Equal(new[] { "1", "2", "II", "005" }, report.Results
                    .Where(result => result.FieldType == WordFieldType.Page)
                    .Select(result => result.ResultText)
                    .ToArray());

                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.NumPages &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "5");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.NumPages &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "005" &&
                    result.Message.Contains("Numeric picture", StringComparison.Ordinal));
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.NumPages &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "V" &&
                    result.Message.Contains("General numeric format", StringComparison.Ordinal));

                WordFieldUpdateResult unsupported = Assert.Single(report.Results, result => result.FieldType == WordFieldType.Database);
                Assert.Equal(WordFieldUpdateStatus.Unsupported, unsupported.Status);
                Assert.Contains("not evaluated", unsupported.Message, StringComparison.OrdinalIgnoreCase);

                WordFieldUpdateResult parseError = Assert.Single(report.Results, result => result.Status == WordFieldUpdateStatus.ParseError);
                Assert.Null(parseError.FieldType);
                Assert.Contains("couldn't be processed", parseError.Message, StringComparison.OrdinalIgnoreCase);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                Assert.Equal(new[] { "1", "2", "II", "005" }, fields
                    .Where(field => field.FieldType == WordFieldType.Page)
                    .Select(field => field.ResultText)
                    .ToArray());
                Assert.Contains(fields, field => field.FieldType == WordFieldType.NumPages && field.ResultText == "5");
                Assert.Contains(fields, field => field.FieldType == WordFieldType.NumPages && field.ResultText == "005");
                Assert.Contains(fields, field => field.FieldType == WordFieldType.NumPages && field.ResultText == "V");
                Assert.Contains(fields, field => field.FieldType == WordFieldType.Database);
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_AppliesFileNameTextFormatSwitches() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.FileNameFormat.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Upper file: ")._paragraph.Append(BuildSimpleField(" FILENAME \\* Upper ", "stale-upper"));
                document.AddParagraph("Lower path: ")._paragraph.Append(BuildSimpleField(" FILENAME \\p \\* Lower ", "stale-lower"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(2, report.TotalCount);
                Assert.Equal(2, report.UpdatedCount);
                string[] results = report.Results
                    .Where(result => result.FieldType == WordFieldType.FileName && result.Status == WordFieldUpdateStatus.Updated)
                    .Select(result => result.ResultText ?? string.Empty)
                    .ToArray();
                Assert.Contains(Path.GetFileName(filePath).ToUpperInvariant(), results);
                Assert.Contains(filePath.ToLowerInvariant(), results);
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesSectionFields() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.Sections.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("First section: ").AddField(WordFieldType.Section);
                document.AddPageBreak();
                document.AddParagraph("First section pages: ").AddField(WordFieldType.SectionPages);

                WordSection secondSection = document.AddSection(SectionMarkValues.NextPage);
                secondSection.AddParagraph("Second section: ").AddField(WordFieldType.Section);
                secondSection.AddParagraph("Second section formatted: ")._paragraph.Append(BuildSimpleField(" SECTION \\* Roman ", "stale-section-roman"));
                secondSection.AddParagraph()._paragraph.Append(new Run(new Break { Type = BreakValues.Page }));
                secondSection.AddParagraph("Second section pages: ")._paragraph.Append(BuildSimpleField(" SECTIONPAGES \\# \"000\" ", "stale-section-pages"));

                WordSection thirdSection = document.AddSection(SectionMarkValues.Continuous);
                thirdSection.AddParagraph("Third section: ").AddField(WordFieldType.Section);
                thirdSection.AddParagraph("Third section pages: ").AddField(WordFieldType.SectionPages);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(7, report.TotalCount);
                Assert.Equal(7, report.UpdatedCount);
                Assert.Equal(0, report.SkippedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);

                Assert.Equal(new[] { "1", "2", "II", "3" }, report.Results
                    .Where(result => result.FieldType == WordFieldType.Section)
                    .Select(result => result.ResultText)
                    .ToArray());
                Assert.Equal(new[] { "2", "002", "1" }, report.Results
                    .Where(result => result.FieldType == WordFieldType.SectionPages)
                    .Select(result => result.ResultText)
                    .ToArray());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                Assert.Equal(new[] { "1", "2", "II", "3" }, fields
                    .Where(field => field.FieldType == WordFieldType.Section)
                    .Select(field => field.ResultText)
                    .ToArray());
                Assert.Equal(new[] { "2", "002", "1" }, fields
                    .Where(field => field.FieldType == WordFieldType.SectionPages)
                    .Select(field => field.ResultText)
                    .ToArray());
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesFileSizeFields() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.FileSize.docx");
            string expectedBytes = string.Empty;
            string expectedKilobytes = string.Empty;
            string expectedMegabytes = string.Empty;
            string expectedFormattedKilobytes = string.Empty;
            string expectedRomanKilobytes = string.Empty;

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Default size: ").AddField(WordFieldType.FileSize);
                document.AddParagraph("Kilobytes: ")._paragraph.Append(BuildSimpleField(" FILESIZE \\k ", "stale-k"));
                document.AddParagraph("Megabytes: ")._paragraph.Append(BuildSimpleField(" FILESIZE \\m ", "stale-m"));
                document.AddParagraph("Formatted kilobytes: ")._paragraph.Append(BuildSimpleField(" FILESIZE \\k \\# \"000\" ", "stale-picture"));
                document.AddParagraph("Roman kilobytes: ")._paragraph.Append(BuildSimpleField(" FILESIZE \\k \\* Roman ", "stale-roman"));
                document.AddParagraph("Unsupported switch: ")._paragraph.Append(BuildSimpleField(" FILESIZE \\p ", "stale-switch"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                long bytes = new FileInfo(filePath).Length;
                expectedBytes = bytes.ToString(System.Globalization.CultureInfo.InvariantCulture);
                expectedKilobytes = Math.Round(bytes / 1000m, 0, MidpointRounding.AwayFromZero).ToString(System.Globalization.CultureInfo.InvariantCulture);
                expectedMegabytes = Math.Round(bytes / 1000000m, 0, MidpointRounding.AwayFromZero).ToString(System.Globalization.CultureInfo.InvariantCulture);
                expectedFormattedKilobytes = Math.Round(bytes / 1000m, 0, MidpointRounding.AwayFromZero).ToString("000", System.Globalization.CultureInfo.InvariantCulture);
                expectedRomanKilobytes = ToRoman((int)Math.Round(bytes / 1000m, 0, MidpointRounding.AwayFromZero));

                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(6, report.TotalCount);
                Assert.Equal(5, report.UpdatedCount);
                Assert.Equal(1, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);

                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.FileSize &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.InstructionText.Contains("FILESIZE", StringComparison.Ordinal) &&
                    !result.InstructionText.Contains("\\k", StringComparison.Ordinal) &&
                    !result.InstructionText.Contains("\\m", StringComparison.Ordinal) &&
                    result.ResultText == expectedBytes);
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.FileSize &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.InstructionText.Contains("\\k", StringComparison.Ordinal) &&
                    !result.InstructionText.Contains("\\#", StringComparison.Ordinal) &&
                    result.ResultText == expectedKilobytes);
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.FileSize &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.InstructionText.Contains("\\m", StringComparison.Ordinal) &&
                    result.ResultText == expectedMegabytes);
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.FileSize &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.InstructionText.Contains("\\#", StringComparison.Ordinal) &&
                    result.ResultText == expectedFormattedKilobytes);
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.FileSize &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.InstructionText.Contains("\\* Roman", StringComparison.Ordinal) &&
                    result.ResultText == expectedRomanKilobytes);

                WordFieldUpdateResult unsupported = Assert.Single(report.Results, result =>
                    result.FieldType == WordFieldType.FileSize &&
                    result.Status == WordFieldUpdateStatus.Unsupported);
                Assert.Contains("\\p", unsupported.Message, StringComparison.Ordinal);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.FileSize &&
                    !field.InstructionText.Contains("\\k", StringComparison.Ordinal) &&
                    !field.InstructionText.Contains("\\m", StringComparison.Ordinal) &&
                    field.ResultText == expectedBytes);
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.FileSize &&
                    field.InstructionText.Contains("\\k", StringComparison.Ordinal) &&
                    !field.InstructionText.Contains("\\#", StringComparison.Ordinal) &&
                    field.ResultText == expectedKilobytes);
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.FileSize &&
                    field.InstructionText.Contains("\\m", StringComparison.Ordinal) &&
                    field.ResultText == expectedMegabytes);
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.FileSize &&
                    field.InstructionText.Contains("\\#", StringComparison.Ordinal) &&
                    field.ResultText == expectedFormattedKilobytes);
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.FileSize &&
                    field.InstructionText.Contains("\\* Roman", StringComparison.Ordinal) &&
                    field.ResultText == expectedRomanKilobytes);
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.FileSize &&
                    field.InstructionText.Contains("\\p", StringComparison.Ordinal) &&
                    field.ResultText == "stale-switch");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesDocumentStatisticsFields() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.DocumentStatistics.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.ApplicationProperties.Words = new DocumentFormat.OpenXml.ExtendedProperties.Words { Text = "42" };
                document.ApplicationProperties.Characters = "1234";

                document.AddParagraph("Words: ").AddField(WordFieldType.NumWords);
                document.AddParagraph("Characters: ").AddField(WordFieldType.NumChars);
                document.AddParagraph("Padded words: ")._paragraph.Append(BuildSimpleField(" NUMWORDS \\# \"000\" ", "stale-padded-words"));
                document.AddParagraph("Grouped characters: ")._paragraph.Append(BuildSimpleField(" NUMCHARS \\# \"#,##0\" ", "stale-grouped-characters"));
                document.AddParagraph("Roman words: ")._paragraph.Append(BuildSimpleField(" NUMWORDS \\* Roman ", "stale-roman-words"));
                document.AddParagraph("Roman characters: ")._paragraph.Append(BuildSimpleField(" NUMCHARS \\* Roman ", "stale-roman-characters"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(6, report.TotalCount);
                Assert.Equal(6, report.UpdatedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);

                Assert.Equal(new[] { "42", "042", "XLII" }, report.Results
                    .Where(result => result.FieldType == WordFieldType.NumWords)
                    .Select(result => result.ResultText)
                    .ToArray());
                Assert.Equal(new[] { "1234", "1,234", "MCCXXXIV" }, report.Results
                    .Where(result => result.FieldType == WordFieldType.NumChars)
                    .Select(result => result.ResultText)
                    .ToArray());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                Assert.Contains(fields, field => field.FieldType == WordFieldType.NumWords && field.ResultText == "42");
                Assert.Contains(fields, field => field.FieldType == WordFieldType.NumWords && field.ResultText == "042");
                Assert.Contains(fields, field => field.FieldType == WordFieldType.NumWords && field.ResultText == "XLII");
                Assert.Contains(fields, field => field.FieldType == WordFieldType.NumChars && field.ResultText == "1234");
                Assert.Contains(fields, field => field.FieldType == WordFieldType.NumChars && field.ResultText == "1,234");
                Assert.Contains(fields, field => field.FieldType == WordFieldType.NumChars && field.ResultText == "MCCXXXIV");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesDateTimeFieldsWithOptionsAndCustomFormats() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.DateTime.docx");
            var updateTime = new DateTime(2026, 6, 30, 15, 16, 17);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Created = new DateTime(2024, 1, 2, 3, 4, 5);
                document.BuiltinDocumentProperties.Modified = new DateTime(2024, 1, 3, 4, 5, 6);
                document.BuiltinDocumentProperties.LastPrinted = new DateTime(2024, 1, 4, 5, 6, 7);

                document.AddParagraph("Created custom: ").AddField(WordFieldType.CreateDate, customFormat: "yyyy/MM/dd");
                document.AddParagraph("Saved default: ").AddField(WordFieldType.SaveDate);
                document.AddParagraph("Printed custom: ").AddField(WordFieldType.PrintDate, customFormat: "yyyy-MM-dd HH:mm");
                document.AddParagraph("Date custom: ").AddField(WordFieldType.Date, customFormat: "yyyy-MM-dd");
                document.AddParagraph("Time custom: ").AddField(WordFieldType.Time, customFormat: "HH:mm");
                document.AddParagraph("Time ampm: ").AddField(WordFieldType.Time, customFormat: "h:mm am/pm");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport(new WordFieldUpdateOptions {
                    CurrentDateTime = updateTime
                });

                Assert.Equal(6, report.TotalCount);
                Assert.Equal(6, report.UpdatedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);

                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.CreateDate &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "2024/01/02");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.SaveDate &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "2024-01-03 04:05:06");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.PrintDate &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "2024-01-04 05:06");
                Assert.Equal(new[] { "2026-06-30" }, report.Results
                    .Where(result => result.FieldType == WordFieldType.Date)
                    .Select(result => result.ResultText)
                    .ToArray());
                Assert.Equal(new[] { "15:16", "3:16 PM" }, report.Results
                    .Where(result => result.FieldType == WordFieldType.Time)
                    .Select(result => result.ResultText)
                    .ToArray());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.CreateDate &&
                    field.Switches.Contains("\\@ \"yyyy/MM/dd\"") &&
                    field.ResultText == "2024/01/02");
                Assert.Contains(fields, field => field.FieldType == WordFieldType.PrintDate && field.ResultText == "2024-01-04 05:06");
                Assert.Contains(fields, field => field.FieldType == WordFieldType.Date && field.ResultText == "2026-06-30");
                Assert.Contains(fields, field => field.FieldType == WordFieldType.Time && field.ResultText == "15:16");
                Assert.Contains(fields, field => field.FieldType == WordFieldType.Time && field.ResultText == "3:16 PM");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_AppliesTextFormatSwitchesToDateFields() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.DateTextFormat.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Created = new DateTime(2024, 1, 2, 3, 4, 5);
                document.AddParagraph("Created upper: ")._paragraph.Append(BuildSimpleField(" CREATEDATE \\@ \"MMMM d, yyyy\" \\* Upper ", "stale-created"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                AssertUpdated(report, WordFieldType.CreateDate, "JANUARY 2, 2024");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.CreateDate &&
                    result.Message.Contains("Text format switch was applied", StringComparison.Ordinal));
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesInfoAndRevisionMetadataFields() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.InfoRevision.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Title = "Premium Gap Plan";
                document.BuiltinDocumentProperties.Category = "Market readiness";
                document.BuiltinDocumentProperties.Revision = "7";
                document.BuiltinDocumentProperties.Version = "2026.6";
                document.BuiltinDocumentProperties.LastPrinted = new DateTime(2024, 1, 4, 5, 6, 7);

                document.AddParagraph("Revision direct: ").AddField(WordFieldType.RevNum);
                document.AddParagraph("Info title: ").AddField(WordFieldType.Info, parameters: new List<string> { "Title" });
                document.AddParagraph("Info category: ").AddField(WordFieldType.Info, parameters: new List<string> { "Category" });
                document.AddParagraph("Info revision: ").AddField(WordFieldType.Info, parameters: new List<string> { "Revision" });
                document.AddParagraph("Info version: ").AddField(WordFieldType.Info, parameters: new List<string> { "Version" });
                document.AddParagraph("Info printed: ").AddField(WordFieldType.Info, parameters: new List<string> { "LastPrinted" });
                document.AddParagraph("Info printed custom: ")._paragraph.Append(BuildSimpleField(" INFO LastPrinted \\@ \"yyyy/MM/dd\" ", "stale-info-printed"));
                document.AddParagraph("Property category: ").AddField(WordFieldType.DocProperty, parameters: new List<string> { "\"Category\"" });
                document.AddParagraph("Property version: ").AddField(WordFieldType.DocProperty, parameters: new List<string> { "\"Version\"" });
                document.AddParagraph("Missing info: ")._paragraph.Append(BuildSimpleField(" INFO MissingBuiltInProperty ", "stale-missing-info"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(10, report.TotalCount);
                Assert.Equal(9, report.UpdatedCount);
                Assert.Equal(1, report.SkippedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);

                AssertUpdated(report, WordFieldType.RevNum, "7");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Info &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.InstructionText.Contains("Title", StringComparison.Ordinal) &&
                    result.ResultText == "Premium Gap Plan");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Info &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.InstructionText.Contains("Category", StringComparison.Ordinal) &&
                    result.ResultText == "Market readiness");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Info &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.InstructionText.Contains("Revision", StringComparison.Ordinal) &&
                    result.ResultText == "7");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Info &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.InstructionText.Contains("Version", StringComparison.Ordinal) &&
                    result.ResultText == "2026.6");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Info &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.InstructionText.Contains("LastPrinted", StringComparison.Ordinal) &&
                    result.ResultText == "2024-01-04 05:06:07");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Info &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.InstructionText.Contains("\\@ \"yyyy/MM/dd\"", StringComparison.Ordinal) &&
                    result.ResultText == "2024/01/04");
                AssertDocPropertyUpdated(report, "Category", "Market readiness");
                AssertDocPropertyUpdated(report, "Version", "2026.6");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Info &&
                    result.Status == WordFieldUpdateStatus.Skipped &&
                    result.InstructionText.Contains("MissingBuiltInProperty", StringComparison.Ordinal) &&
                    result.ResultText == null);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                Assert.Contains(fields, field => field.FieldType == WordFieldType.RevNum && field.ResultText == "7");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Info &&
                    field.InstructionText.Contains("Category", StringComparison.Ordinal) &&
                    field.ResultText == "Market readiness");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Info &&
                    field.InstructionText.Contains("LastPrinted", StringComparison.Ordinal) &&
                    field.ResultText == "2024-01-04 05:06:07");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.DocProperty &&
                    field.InstructionText.Contains("Version", StringComparison.Ordinal) &&
                    field.ResultText == "2026.6");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Info &&
                    field.InstructionText.Contains("MissingBuiltInProperty", StringComparison.Ordinal) &&
                    field.ResultText == "stale-missing-info");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_ReportsUnsupportedDiagnosticsForKnownFieldFormatSwitches() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.UnsupportedKnownFormatDiagnostics.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Target heading").AddBookmark("TargetBookmark");
                document.AddParagraph("Unknown named format: ")._paragraph.Append(BuildSimpleField(" REF TargetBookmark \\* FutureCase ", "stale-ref"));
                document.AddParagraph("Unsupported numeric picture: ")._paragraph.Append(BuildSimpleField(" AUTHOR \\# \"000\" ", "stale-author"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(2, report.TotalCount);
                Assert.Equal(0, report.UpdatedCount);
                Assert.Equal(2, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);

                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Ref &&
                    result.Status == WordFieldUpdateStatus.Unsupported &&
                    result.Message.Contains("FutureCase", StringComparison.Ordinal));
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Author &&
                    result.Status == WordFieldUpdateStatus.Unsupported &&
                    result.Message.Contains("numeric picture", StringComparison.OrdinalIgnoreCase));

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Ref &&
                    field.UnsupportedParseDetails.Any(detail => detail.Contains("FutureCase", StringComparison.Ordinal)) &&
                    field.ResultText == "stale-ref");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Author &&
                    field.UnsupportedParseDetails.Any(detail => detail.Contains("numeric picture", StringComparison.OrdinalIgnoreCase)) &&
                    field.ResultText == "stale-author");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_PreservesRefBookmarkRangeSeparators() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.RefBookmarkRangeSeparators.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document._document.Body!.Append(
                    new Paragraph(
                        new BookmarkStart { Name = "ParagraphSpan", Id = "310" },
                        new Run(new Text("First paragraph") { Space = SpaceProcessingModeValues.Preserve })),
                    new Paragraph(
                        new Run(new Text("Second paragraph") { Space = SpaceProcessingModeValues.Preserve }),
                        new BookmarkEnd { Id = "310" }),
                    new Table(
                        new TableRow(
                            new TableCell(
                                new Paragraph(
                                    new BookmarkStart { Name = "CellSpan", Id = "311" },
                                    new Run(new Text("First cell") { Space = SpaceProcessingModeValues.Preserve }))),
                            new TableCell(
                                new Paragraph(
                                    new Run(new Text("Second cell") { Space = SpaceProcessingModeValues.Preserve }),
                                    new BookmarkEnd { Id = "311" })))),
                    new Paragraph(new Run(new Text("Paragraph REF: ") { Space = SpaceProcessingModeValues.Preserve }), BuildSimpleField(" REF ParagraphSpan ", "stale-paragraph")),
                    new Paragraph(new Run(new Text("Cell REF: ") { Space = SpaceProcessingModeValues.Preserve }), BuildSimpleField(" REF CellSpan ", "stale-cell")));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Ref &&
                    result.ResultText == "First paragraph\nSecond paragraph");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Ref &&
                    result.ResultText == "First cell\tSecond cell");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesLiteralQuoteFields() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.QuoteLiterals.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Simple quote: ")._paragraph.Append(BuildSimpleField(" QUOTE \"Automation ready\" ", "stale-simple"));
                AddComplexField(document.AddParagraph("Complex quote: ")._paragraph, "stale-complex", " QUOTE ", "\"Complex literal\" ");
                document.AddParagraph("Container quote: ")._paragraph.Append(BuildSimpleField(" QUOTE ", "stale-empty"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(3, report.TotalCount);
                Assert.Equal(2, report.UpdatedCount);
                Assert.Equal(1, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);

                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Quote &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "Automation ready");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Quote &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "Complex literal");

                WordFieldUpdateResult unsupported = Assert.Single(report.Results, result =>
                    result.FieldType == WordFieldType.Quote &&
                    result.Status == WordFieldUpdateStatus.Unsupported);
                Assert.Contains("quoted literal", unsupported.Message, StringComparison.OrdinalIgnoreCase);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Quote &&
                    field.InstructionText.Contains("Automation ready", StringComparison.Ordinal) &&
                    field.ResultText == "Automation ready");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Quote &&
                    field.InstructionText.Contains("Complex literal", StringComparison.Ordinal) &&
                    field.ResultText == "Complex literal");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Quote &&
                    string.Equals(field.InstructionText.Trim(), "QUOTE", StringComparison.OrdinalIgnoreCase) &&
                    field.ResultText == "stale-empty");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_AppliesDeterministicQuoteFormatSwitches() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.QuoteFormats.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Upper quote: ")._paragraph.Append(BuildSimpleField(" QUOTE \"mixed status\" \\* Upper ", "stale-upper"));
                AddComplexField(document.AddParagraph("Caps quote: ")._paragraph, "stale-caps", " QUOTE ", "\"network status\" ", "\\* Caps ");
                document.AddParagraph("Roman quote: ")._paragraph.Append(BuildSimpleField(" QUOTE \"12\" \\* Roman ", "stale-roman-number"));
                document.AddParagraph("Lower roman quote: ")._paragraph.Append(BuildSimpleField(" QUOTE \"12\" \\* roman ", "stale-lower-roman-number"));
                document.AddParagraph("Alphabetic quote: ")._paragraph.Append(BuildSimpleField(" QUOTE \"27\" \\* ALPHABETICAL ", "stale-alpha-number"));
                document.AddParagraph("Ordinal quote: ")._paragraph.Append(BuildSimpleField(" QUOTE \"21\" \\* Ordinal ", "stale-ordinal-number"));
                document.AddParagraph("Hex quote: ")._paragraph.Append(BuildSimpleField(" QUOTE \"255\" \\* Hex ", "stale-hex-number"));
                document.AddParagraph("Cardinal text quote: ")._paragraph.Append(BuildSimpleField(" QUOTE \"21\" \\* CardText ", "stale-card-text"));
                document.AddParagraph("Ordinal text quote: ")._paragraph.Append(BuildSimpleField(" QUOTE \"21\" \\* OrdText ", "stale-ord-text"));
                document.AddParagraph("Dollar text quote: ")._paragraph.Append(BuildSimpleField(" QUOTE \"12\" \\* DollarText ", "stale-dollar-text"));
                document.AddParagraph("Numeric picture quote: ")._paragraph.Append(BuildSimpleField(" QUOTE \"1234.5\" \\# \"#,##0.00\" ", "stale-picture"));
                document.AddParagraph("Unsupported quote: ")._paragraph.Append(BuildSimpleField(" QUOTE \"seven\" \\* Roman ", "stale-roman"));
                document.AddParagraph("Unsupported negative hex quote: ")._paragraph.Append(BuildSimpleField(" QUOTE \"-1\" \\* Hex ", "stale-negative-hex"));
                document.AddParagraph("Unsupported negative dollar quote: ")._paragraph.Append(BuildSimpleField(" QUOTE \"-1\" \\* DollarText ", "stale-negative-dollar"));
                document.AddParagraph("Unsupported mixed quote formats: ")._paragraph.Append(BuildSimpleField(" QUOTE \"1234\" \\* Roman \\# \"0000\" ", "stale-mixed-picture-format"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(15, report.TotalCount);
                Assert.Equal(11, report.UpdatedCount);
                Assert.Equal(4, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);

                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Quote &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "MIXED STATUS");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Quote &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "Network Status");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Quote &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "XII");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Quote &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "xii");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Quote &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "AA");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Quote &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "21st");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Quote &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "FF");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Quote &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "twenty-one");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Quote &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "twenty-first");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Quote &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "twelve and 00/100");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Quote &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "1,234.50");

                WordFieldUpdateResult[] unsupported = report.Results.Where(result =>
                    result.FieldType == WordFieldType.Quote &&
                    result.Status == WordFieldUpdateStatus.Unsupported).ToArray();
                Assert.Equal(4, unsupported.Length);
                Assert.Contains(unsupported, result => result.Message.Contains("Roman", StringComparison.Ordinal));
                Assert.Contains(unsupported, result => result.Message.Contains("Hex", StringComparison.Ordinal));
                Assert.Contains(unsupported, result => result.Message.Contains("DollarText", StringComparison.Ordinal));
                Assert.Contains(unsupported, result => result.Message.Contains("combine", StringComparison.OrdinalIgnoreCase));

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Quote &&
                    field.FormatSwitches.Contains(WordFieldFormat.Upper) &&
                    field.ResultText == "MIXED STATUS");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Quote &&
                    field.FormatSwitches.Contains(WordFieldFormat.Caps) &&
                    field.ResultText == "Network Status");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Quote &&
                    field.FormatSwitches.Contains(WordFieldFormat.Roman) &&
                    field.ResultText == "XII");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Quote &&
                    field.FormatSwitches.Contains(WordFieldFormat.roman) &&
                    field.ResultText == "xii");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Quote &&
                    field.FormatSwitches.Contains(WordFieldFormat.ALPHABETICAL) &&
                    field.ResultText == "AA");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Quote &&
                    field.FormatSwitches.Contains(WordFieldFormat.Ordinal) &&
                    field.ResultText == "21st");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Quote &&
                    field.FormatSwitches.Contains(WordFieldFormat.Hex) &&
                    field.ResultText == "FF");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Quote &&
                    field.FormatSwitches.Contains(WordFieldFormat.CardText) &&
                    field.ResultText == "twenty-one");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Quote &&
                    field.FormatSwitches.Contains(WordFieldFormat.OrdText) &&
                    field.ResultText == "twenty-first");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Quote &&
                    field.FormatSwitches.Contains(WordFieldFormat.DollarText) &&
                    field.ResultText == "twelve and 00/100");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Quote &&
                    field.InstructionText.Contains(@"\# ""#,##0.00""", StringComparison.Ordinal) &&
                    field.ResultText == "1,234.50");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Quote &&
                    field.FormatSwitches.Contains(WordFieldFormat.Roman) &&
                    field.ResultText == "stale-roman");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Quote &&
                    field.FormatSwitches.Contains(WordFieldFormat.Hex) &&
                    field.ResultText == "stale-negative-hex");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Quote &&
                    field.FormatSwitches.Contains(WordFieldFormat.DollarText) &&
                    field.ResultText == "stale-negative-dollar");
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Quote &&
                    field.InstructionText.Contains(@"\# ""0000""", StringComparison.Ordinal) &&
                    field.FormatSwitches.Contains(WordFieldFormat.Roman) &&
                    field.ResultText == "stale-mixed-picture-format");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_AppliesTextFormatSwitchesToMetadataFields() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.MetadataTextFormats.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Creator = "Ada Lovelace";
                document.BuiltinDocumentProperties.Title = "MIXED TITLE";
                document.BuiltinDocumentProperties.Subject = "release status";
                document.AddParagraph("Author upper: ")._paragraph.Append(BuildSimpleField(" AUTHOR \\* Upper ", "stale-author"));
                document.AddParagraph("Title lower: ")._paragraph.Append(BuildSimpleField(" TITLE \\* Lower ", "stale-title"));
                document.AddParagraph("Subject caps: ")._paragraph.Append(BuildSimpleField(" SUBJECT \\* Caps ", "stale-subject"));
                document.AddParagraph("Unsupported author: ")._paragraph.Append(BuildSimpleField(" AUTHOR \\* Roman ", "stale-roman-author"));
                document.AddParagraph("Unsupported property: ")._paragraph.Append(BuildSimpleField(" DOCPROPERTY Author \\* Roman ", "stale-roman-property"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(5, report.TotalCount);
                Assert.Equal(3, report.UpdatedCount);
                Assert.Equal(2, report.UnsupportedCount);
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Author &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "ADA LOVELACE");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Title &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "mixed title");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Subject &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "Release Status");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Author &&
                    result.Status == WordFieldUpdateStatus.Unsupported &&
                    result.Message.Contains("Roman", StringComparison.Ordinal));
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.DocProperty &&
                    result.Status == WordFieldUpdateStatus.Unsupported &&
                    result.Message.Contains("Roman", StringComparison.Ordinal));
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_HandlesNestedComplexFieldsWithoutCorruptingContainingResults() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.NestedComplexFields.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Creator = "Ada Lovelace";
                document.BuiltinDocumentProperties.Title = "Premium Gap Plan";

                AddNestedComplexFields(
                    document.AddParagraph()._paragraph,
                    " QUOTE ",
                    " AUTHOR ",
                    "Outer start ",
                    "Nested Author stale",
                    " outer end");
                AddNestedComplexFields(
                    document.AddParagraph()._paragraph,
                    " AUTHOR ",
                    " TITLE ",
                    "Outer author stale ",
                    "Nested Title stale",
                    " author tail");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(4, report.TotalCount);
                Assert.Equal(2, report.UpdatedCount);
                Assert.Equal(1, report.UnsupportedCount);
                Assert.Equal(1, report.SkippedCount);
                Assert.Equal(0, report.ParseErrorCount);

                WordFieldUpdateResult quote = Assert.Single(report.Results, result => result.FieldType == WordFieldType.Quote);
                Assert.Equal(WordFieldUpdateStatus.Unsupported, quote.Status);

                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Author &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "Ada Lovelace");

                WordFieldUpdateResult nestedTitle = Assert.Single(report.Results, result =>
                    result.FieldType == WordFieldType.Title);
                Assert.Equal(WordFieldUpdateStatus.Skipped, nestedTitle.Status);
                Assert.Contains("containing field", nestedTitle.Message, StringComparison.OrdinalIgnoreCase);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                WordFieldInfo quote = Assert.Single(fields, field => field.FieldType == WordFieldType.Quote);
                Assert.Equal("Outer start Ada Lovelace outer end", quote.ResultText);

                WordFieldInfo nestedAuthor = Assert.Single(fields, field =>
                    field.FieldType == WordFieldType.Author &&
                    field.NestingLevel == 1);
                Assert.Equal("Ada Lovelace", nestedAuthor.ResultText);

                WordFieldInfo outerAuthor = Assert.Single(fields, field =>
                    field.FieldType == WordFieldType.Author &&
                    field.NestingLevel == 0);
                Assert.Equal("Ada Lovelace", outerAuthor.ResultText);

                WordFieldInfo nestedTitle = Assert.Single(fields, field => field.FieldType == WordFieldType.Title);
                Assert.Equal(string.Empty, nestedTitle.ResultText);
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_KeepsEndRunTrailingTextOutOfComplexResult() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.ComplexEndRunTrailingText.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                Paragraph paragraph = document.AddParagraph("Value: ")._paragraph;
                paragraph.Append(
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode(" QUOTE \"updated\" ") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                    new Run(new Text("stale") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(
                        new FieldChar { FieldCharType = FieldCharValues.End },
                        new Text(" trailing") { Space = SpaceProcessingModeValues.Preserve }));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldInfo inspectedField = Assert.Single(document.InspectFields(), field => field.FieldType == WordFieldType.Quote);
                Assert.Equal("stale", inspectedField.ResultText);
                Assert.DoesNotContain(" trailing", inspectedField.ResultText, StringComparison.Ordinal);

                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Contains(report.Results, field =>
                    field.FieldType == WordFieldType.Quote &&
                    field.Status == WordFieldUpdateStatus.Updated &&
                    field.ResultText == "updated");
                Assert.DoesNotContain(report.Results, field =>
                    field.FieldType == WordFieldType.Quote &&
                    field.ResultText.Contains(" trailing", StringComparison.Ordinal));
                Assert.Contains(document._document.Body!.Descendants<Text>(), text => text.Text == "updated");
                Assert.Contains(document._document.Body!.Descendants<Text>(), text => text.Text == " trailing");
                Assert.DoesNotContain(document._document.Body!.Descendants<Text>(), text => text.Text == "stale");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_PreservesEndRunSuffixWhenResultSharesEndRun() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.ComplexResultSharesEndRun.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                Paragraph paragraph = document.AddParagraph("Value: ")._paragraph;
                paragraph.Append(
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                    new Run(new FieldCode(" QUOTE \"fresh\" ") { Space = SpaceProcessingModeValues.Preserve }),
                    new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                    new Run(
                        new Text("stale") { Space = SpaceProcessingModeValues.Preserve },
                        new FieldChar { FieldCharType = FieldCharValues.End },
                        new Text(" suffix") { Space = SpaceProcessingModeValues.Preserve }));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                WordFieldUpdateResult result = Assert.Single(report.Results, field => field.FieldType == WordFieldType.Quote);
                Assert.Equal("fresh", result.ResultText);
                Assert.Contains(document._document.Body!.Descendants<Text>(), text => text.Text == "fresh");
                Assert.Contains(document._document.Body!.Descendants<Text>(), text => text.Text == " suffix");
                Assert.DoesNotContain(document._document.Body!.Descendants<Text>(), text => text.Text == "stale");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesSameRunComplexFieldMarkers() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.SameRunComplexFieldMarkers.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                Paragraph paragraph = document.AddParagraph("Value: ")._paragraph;
                paragraph.Append(new Run(
                    new FieldChar { FieldCharType = FieldCharValues.Begin },
                    new FieldCode(" QUOTE \"fresh\" ") { Space = SpaceProcessingModeValues.Preserve },
                    new FieldChar { FieldCharType = FieldCharValues.Separate },
                    new Text("stale") { Space = SpaceProcessingModeValues.Preserve },
                    new FieldChar { FieldCharType = FieldCharValues.End }));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                WordFieldUpdateResult result = Assert.Single(report.Results, field => field.FieldType == WordFieldType.Quote);
                Assert.Equal(WordFieldUpdateStatus.Updated, result.Status);
                Assert.Equal("fresh", result.ResultText);
                Assert.Equal("Value: fresh", string.Concat(document._document.Body!.Descendants<Text>().Select(text => text.Text)));
                Assert.DoesNotContain(document._document.Body!.Descendants<Text>(), text => text.Text == "stale");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_PreservesSwitchLookingTextInsideQuoteLiteral() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.QuoteLiteralWithSwitchText.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph()._paragraph.Append(BuildSimpleField("QUOTE \"Use \\* Upper here\"", "stale"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                WordFieldUpdateResult result = Assert.Single(report.Results, field => field.FieldType == WordFieldType.Quote);
                Assert.Equal(WordFieldUpdateStatus.Updated, result.Status);
                Assert.Equal("Use \\* Upper here", result.ResultText);
                Assert.Contains(document._document.Body!.Descendants<Text>(), text => text.Text == "Use \\* Upper here");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesNestedInstructionComplexFieldsBeforeContainingFormula() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.NestedInstructionFormula.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.CustomDocumentProperties["Rate"] = new WordCustomProperty("0.125");
                AddNestedInstructionComplexField(
                    document.AddParagraph()._paragraph,
                    " = ",
                    " DOCPROPERTY Rate ",
                    "stale-rate",
                    " * 100 \\# \"0.0\" ",
                    "stale-formula");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(2, report.TotalCount);
                Assert.Equal(2, report.UpdatedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.SkippedCount);
                Assert.Equal(0, report.ParseErrorCount);

                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.DocProperty &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "0.125");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Formula &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "12.5");

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                WordFieldInfo formula = Assert.Single(fields, field => field.FieldType == WordFieldType.Formula);
                Assert.Equal(WordFieldRepresentation.Complex, formula.Representation);
                Assert.Equal("12.5", formula.ResultText);
                Assert.Contains("0.125", formula.InstructionText, StringComparison.Ordinal);

                WordFieldInfo nestedProperty = Assert.Single(fields, field => field.FieldType == WordFieldType.DocProperty);
                Assert.Equal(1, nestedProperty.NestingLevel);
                Assert.Equal("0.125", nestedProperty.ResultText);
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesRefAndPageRefFields() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.References.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Target heading").AddBookmark("TargetBookmark");
                document.AddPageBreak();
                document.AddParagraph("Reference: ").AddField(WordFieldType.Ref, parameters: new List<string> { "TargetBookmark" });
                document.AddParagraph("Page reference: ").AddField(WordFieldType.PageRef, parameters: new List<string> { "TargetBookmark" });
                document.AddParagraph("Missing reference: ").AddField(WordFieldType.Ref, parameters: new List<string> { "MissingBookmark" });
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(3, report.TotalCount);
                Assert.Equal(2, report.UpdatedCount);
                Assert.Equal(1, report.SkippedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);

                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Ref &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "Target heading");
                AssertUpdated(report, WordFieldType.PageRef, "1");

                WordFieldUpdateResult missing = Assert.Single(report.Results, result =>
                    result.FieldType == WordFieldType.Ref &&
                    result.Status == WordFieldUpdateStatus.Skipped);
                Assert.Contains("MissingBookmark", missing.Message, StringComparison.Ordinal);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                Assert.Contains(fields, field => field.FieldType == WordFieldType.Ref && field.ResultText == "Target heading");
                Assert.Contains(fields, field => field.FieldType == WordFieldType.PageRef && field.ResultText == "1");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesRefFieldsForRelatedPartBookmarks() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.RelatedPartReferences.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Body source").AddBookmark("BodyBookmark");
                document.AddHeadersAndFooters();

                Header header = document._wordprocessingDocument.MainDocumentPart!.HeaderParts.Single().Header!;
                Footer footer = document._wordprocessingDocument.MainDocumentPart!.FooterParts.Single().Footer!;
                AppendBookmarkedParagraph(header, "HeaderBookmark", "7201", "Header source");
                AppendBookmarkedParagraph(footer, "FooterBookmark", "7202", "Footer source");

                document.AddParagraph("Footnote reference anchor").AddFootNote("Footnote body");
                document.AddParagraph("Endnote reference anchor").AddEndNote("Endnote body");
                Footnote footnote = document._wordprocessingDocument.MainDocumentPart!.FootnotesPart!.Footnotes!.Elements<Footnote>().First(note => note.Type == null);
                Endnote endnote = document._wordprocessingDocument.MainDocumentPart!.EndnotesPart!.Endnotes!.Elements<Endnote>().First(note => note.Type == null);
                AppendBookmarkedParagraph(footnote, "FootnoteBookmark", "7203", "Footnote source");
                AppendBookmarkedParagraph(endnote, "EndnoteBookmark", "7204", "Endnote source");

                document.AddParagraph("Header reference: ")._paragraph.Append(BuildSimpleField(" REF HeaderBookmark ", "stale-header"));
                document.AddParagraph("Footer reference: ")._paragraph.Append(BuildSimpleField(" REF FooterBookmark ", "stale-footer"));
                document.AddParagraph("Footnote reference: ")._paragraph.Append(BuildSimpleField(" REF FootnoteBookmark ", "stale-footnote"));
                document.AddParagraph("Endnote reference: ")._paragraph.Append(BuildSimpleField(" REF EndnoteBookmark ", "stale-endnote"));
                document.AddParagraph("Related-part page reference: ")._paragraph.Append(BuildSimpleField(" PAGEREF HeaderBookmark ", "stale-page"));
                header.Append(new Paragraph(new Run(new Text("Header field: ") { Space = SpaceProcessingModeValues.Preserve }), BuildSimpleField(" REF FootnoteBookmark ", "stale-header-field")));
                footer.Append(new Paragraph(new Run(new Text("Footer field: ") { Space = SpaceProcessingModeValues.Preserve }), BuildSimpleField(" REF EndnoteBookmark ", "stale-footer-field")));
                footnote.Append(new Paragraph(new Run(new Text("Footnote field: ") { Space = SpaceProcessingModeValues.Preserve }), BuildSimpleField(" REF HeaderBookmark ", "stale-footnote-field")));
                endnote.Append(new Paragraph(new Run(new Text("Endnote field: ") { Space = SpaceProcessingModeValues.Preserve }), BuildSimpleField(" REF FooterBookmark ", "stale-endnote-field")));

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(9, report.TotalCount);
                Assert.Equal(8, report.UpdatedCount);
                Assert.Equal(1, report.SkippedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);

                Assert.Contains(report.Results, result => result.LocationKind == WordFieldLocationKind.Body && result.FieldType == WordFieldType.Ref && result.ResultText == "Header source");
                Assert.Contains(report.Results, result => result.LocationKind == WordFieldLocationKind.Body && result.FieldType == WordFieldType.Ref && result.ResultText == "Footer source");
                Assert.Contains(report.Results, result => result.LocationKind == WordFieldLocationKind.Body && result.FieldType == WordFieldType.Ref && result.ResultText == "Footnote source");
                Assert.Contains(report.Results, result => result.LocationKind == WordFieldLocationKind.Body && result.FieldType == WordFieldType.Ref && result.ResultText == "Endnote source");
                Assert.Contains(report.Results, result => result.LocationKind == WordFieldLocationKind.Header && result.FieldType == WordFieldType.Ref && result.ResultText == "Footnote source");
                Assert.Contains(report.Results, result => result.LocationKind == WordFieldLocationKind.Footer && result.FieldType == WordFieldType.Ref && result.ResultText == "Endnote source");
                Assert.Contains(report.Results, result => result.LocationKind == WordFieldLocationKind.Footnote && result.FieldType == WordFieldType.Ref && result.ResultText == "Header source");
                Assert.Contains(report.Results, result => result.LocationKind == WordFieldLocationKind.Endnote && result.FieldType == WordFieldType.Ref && result.ResultText == "Footer source");

                WordFieldUpdateResult skippedPageRef = Assert.Single(report.Results, result => result.FieldType == WordFieldType.PageRef);
                Assert.Equal(WordFieldUpdateStatus.Skipped, skippedPageRef.Status);
                Assert.Contains("outside the document body", skippedPageRef.Message, StringComparison.OrdinalIgnoreCase);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                Assert.Contains(fields, field => field.LocationKind == WordFieldLocationKind.Body && field.FieldType == WordFieldType.Ref && field.ResultText == "Header source");
                Assert.Contains(fields, field => field.LocationKind == WordFieldLocationKind.Header && field.FieldType == WordFieldType.Ref && field.ResultText == "Footnote source");
                Assert.Contains(fields, field => field.LocationKind == WordFieldLocationKind.Footer && field.FieldType == WordFieldType.Ref && field.ResultText == "Endnote source");
                Assert.Contains(fields, field => field.LocationKind == WordFieldLocationKind.Footnote && field.FieldType == WordFieldType.Ref && field.ResultText == "Header source");
                Assert.Contains(fields, field => field.LocationKind == WordFieldLocationKind.Endnote && field.FieldType == WordFieldType.Ref && field.ResultText == "Footer source");
                Assert.Contains(fields, field => field.FieldType == WordFieldType.PageRef && field.ResultText == "stale-page");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_AppliesDeterministicReferenceFormatSwitches() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.ReferenceFormats.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("network status").AddBookmark("TextBookmark");
                document.AddPageBreak();
                document.AddParagraph("Page target").AddBookmark("PageBookmark");
                for (int pageBreakIndex = 0; pageBreakIndex < 8; pageBreakIndex++) {
                    document.AddPageBreak();
                }

                document.AddParagraph("Hex page target").AddBookmark("HexPageBookmark");
                document.AddParagraph("Upper reference: ").AddField(WordFieldType.Ref, WordFieldFormat.Upper, parameters: new List<string> { "TextBookmark" });
                document.AddParagraph("Lower reference: ").AddField(WordFieldType.Ref, WordFieldFormat.Lower, parameters: new List<string> { "TextBookmark" });
                document.AddParagraph("First cap reference: ").AddField(WordFieldType.Ref, WordFieldFormat.FirstCap, parameters: new List<string> { "TextBookmark" });
                document.AddParagraph("Caps reference: ").AddField(WordFieldType.Ref, WordFieldFormat.Caps, parameters: new List<string> { "TextBookmark" });
                document.AddParagraph("Roman page reference: ").AddField(WordFieldType.PageRef, WordFieldFormat.roman, parameters: new List<string> { "PageBookmark" });
                document.AddParagraph("Alphabetic page reference: ").AddField(WordFieldType.PageRef, WordFieldFormat.ALPHABETICAL, parameters: new List<string> { "PageBookmark" });
                document.AddParagraph("Ordinal page reference: ").AddField(WordFieldType.PageRef, WordFieldFormat.Ordinal, parameters: new List<string> { "PageBookmark" });
                document.AddParagraph("Hex page reference: ").AddField(WordFieldType.PageRef, WordFieldFormat.Hex, parameters: new List<string> { "HexPageBookmark" });
                document.AddParagraph("Cardinal text page reference: ").AddField(WordFieldType.PageRef, WordFieldFormat.CardText, parameters: new List<string> { "PageBookmark" });
                document.AddParagraph("Ordinal text page reference: ").AddField(WordFieldType.PageRef, WordFieldFormat.OrdText, parameters: new List<string> { "PageBookmark" });
                document.AddParagraph("Dollar text page reference: ").AddField(WordFieldType.PageRef, WordFieldFormat.DollarText, parameters: new List<string> { "PageBookmark" });
                document.AddParagraph("Padded page reference: ")._paragraph.Append(BuildSimpleField(" PAGEREF PageBookmark \\# \"000\" ", "stale-pageref-picture"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(12, report.TotalCount);
                Assert.Equal(12, report.UpdatedCount);
                Assert.Equal(0, report.UnsupportedCount);

                Assert.Equal(
                    new[] { "NETWORK STATUS", "network status", "Network status", "Network Status" },
                    report.Results
                        .Where(result => result.FieldType == WordFieldType.Ref)
                        .Select(result => result.ResultText)
                        .ToArray());
                Assert.Equal(
                    new[] { "ii", "B", "2nd", "A", "two", "second", "two and 00/100", "002" },
                    report.Results
                        .Where(result => result.FieldType == WordFieldType.PageRef)
                        .Select(result => result.ResultText)
                        .ToArray());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                Assert.Equal(
                    new[] { "NETWORK STATUS", "network status", "Network status", "Network Status" },
                    fields.Where(field => field.FieldType == WordFieldType.Ref)
                        .Select(field => field.ResultText)
                        .ToArray());
                Assert.Equal(
                    new[] { "ii", "B", "2nd", "A", "two", "second", "two and 00/100", "002" },
                    fields.Where(field => field.FieldType == WordFieldType.PageRef)
                        .Select(field => field.ResultText)
                        .ToArray());
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_CountsParagraphPageStartsForPageReferences() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.PageReferenceParagraphStarts.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("First page");
                document.AddParagraph("Two explicit page breaks")._paragraph.Append(
                    new Run(new Break { Type = BreakValues.Page }),
                    new Run(new Break { Type = BreakValues.Page }));
                Paragraph pageStart = document.AddParagraph("Page-start bookmark target")._paragraph;
                pageStart.ParagraphProperties = new ParagraphProperties(new PageBreakBefore());
                pageStart.InsertBefore(new BookmarkStart { Name = "PageStartBookmark", Id = "101" }, pageStart.GetFirstChild<Run>());
                pageStart.Append(new BookmarkEnd { Id = "101" });
                pageStart.Append(new Run(new Text(" Current page: ") { Space = SpaceProcessingModeValues.Preserve }));
                pageStart.Append(BuildSimpleField(" PAGE ", "stale-page"));
                document.AddParagraph("Page reference: ")._paragraph.Append(BuildSimpleField(" PAGEREF PageStartBookmark ", "stale-pageref"));
                Paragraph inlineBreak = document.AddParagraph("Inline break target")._paragraph;
                inlineBreak.Append(
                    new Run(new Break { Type = BreakValues.Page }),
                    new Run(new Text(" After break ") { Space = SpaceProcessingModeValues.Preserve }),
                    new BookmarkStart { Name = "InlineBreakBookmark", Id = "102" },
                    new Run(new Text("bookmarked") { Space = SpaceProcessingModeValues.Preserve }),
                    new BookmarkEnd { Id = "102" },
                    new Run(new Text(" Current page: ") { Space = SpaceProcessingModeValues.Preserve }),
                    BuildSimpleField(" PAGE ", "stale-inline-page"));
                document.AddParagraph("Inline page reference: ")._paragraph.Append(BuildSimpleField(" PAGEREF InlineBreakBookmark ", "stale-inline-pageref"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(4, report.TotalCount);
                Assert.Equal(4, report.UpdatedCount);
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Page &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "4");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.PageRef &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "4");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Page &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "5");
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.PageRef &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "5");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_ReportsUnsupportedReferenceFormatSwitches() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.ReferenceUnsupportedFormats.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Target heading").AddBookmark("TextBookmark");
                document.AddParagraph("Unsupported reference: ")._paragraph.Append(BuildSimpleField(" REF TextBookmark \\* Roman ", "stale-ref"));
                document.AddParagraph("Unsupported page reference: ")._paragraph.Append(BuildSimpleField(" PAGEREF TextBookmark \\* Upper ", "stale-page"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(2, report.TotalCount);
                Assert.Equal(0, report.UpdatedCount);
                Assert.Equal(2, report.UnsupportedCount);
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Ref &&
                    result.Status == WordFieldUpdateStatus.Unsupported &&
                    result.Message.Contains("Roman", StringComparison.Ordinal));
                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.PageRef &&
                    result.Status == WordFieldUpdateStatus.Unsupported &&
                    result.Message.Contains("Upper", StringComparison.Ordinal));

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                Assert.Contains(fields, field => field.FieldType == WordFieldType.Ref && field.ResultText == "stale-ref");
                Assert.Contains(fields, field => field.FieldType == WordFieldType.PageRef && field.ResultText == "stale-page");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesRefListNumberSwitches() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.ReferenceListNumbers.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordList list = document.AddCustomList();
                AddReferenceNumberingLevel(list, "%CurrentLevel");
                AddReferenceNumberingLevel(list, "%1.%CurrentLevel");
                AddReferenceNumberingLevel(list, "Section %1.%2.%CurrentLevel");

                list.AddItem("Parent A");
                list.AddItem("Parent B");
                list.AddItem("Child one", 1);
                list.AddItem("Grandchild one", 2).AddBookmark("GrandchildOne");
                WordParagraph childTwo = list.AddItem("Child two cross reference: ", 1);
                childTwo._paragraph.Append(BuildSimpleField(" REF GrandchildOne \\r ", "stale-r"));

                document.AddParagraph("No context: ")._paragraph.Append(BuildSimpleField(" REF GrandchildOne \\n ", "stale-n"));
                document.AddParagraph("Full context: ")._paragraph.Append(BuildSimpleField(" REF GrandchildOne \\w ", "stale-w"));
                document.AddParagraph("Full context without text: ")._paragraph.Append(BuildSimpleField(" REF GrandchildOne \\w \\t ", "stale-wt"));

                WordList bulletList = document.AddList(WordListStyle.Bulleted);
                bulletList.AddItem("Bullet target").AddBookmark("BulletTarget");
                document.AddParagraph("Unsupported bullet reference: ")._paragraph.Append(BuildSimpleField(" REF BulletTarget \\n ", "stale-bullet"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(5, report.TotalCount);
                Assert.Equal(4, report.UpdatedCount);
                Assert.Equal(1, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);

                Assert.Equal(
                    new[] { "1.1", "1", "Section 2.1.1", "2.1.1" },
                    report.Results
                        .Where(result => result.FieldType == WordFieldType.Ref && result.Status == WordFieldUpdateStatus.Updated)
                        .Select(result => result.ResultText)
                        .ToArray());

                WordFieldUpdateResult unsupported = Assert.Single(report.Results, result =>
                    result.FieldType == WordFieldType.Ref &&
                    result.Status == WordFieldUpdateStatus.Unsupported);
                Assert.Contains("does not support numbering format", unsupported.Message, StringComparison.OrdinalIgnoreCase);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields()
                    .Where(field => field.FieldType == WordFieldType.Ref)
                    .Select(field => field.ResultText)
                    .ToArray();

                Assert.Equal(new[] { "1.1", "1", "Section 2.1.1", "2.1.1", "stale-bullet" }, fields);
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UsesNumberingLevelOverridesForRefListNumbers() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.ReferenceListNumbers.LevelOverride.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                int numberingId = AddReferenceNumberingDefinition(document, "%1");
                Numbering numbering = document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering;
                NumberingInstance instance = numbering.Elements<NumberingInstance>().Single(item => item.NumberID?.Value == numberingId);
                instance.Append(new LevelOverride(
                    new Level(
                        new StartNumberingValue { Val = 1 },
                        new NumberingFormat { Val = NumberFormatValues.UpperRoman },
                        new LevelText { Val = "Article %1" }) {
                        LevelIndex = 0
                    }) {
                    LevelIndex = 0
                });

                WordParagraph target = document.AddParagraph("Override target");
                target._paragraph.ParagraphProperties = new ParagraphProperties(
                    new NumberingProperties(
                        new NumberingLevelReference { Val = 0 },
                        new NumberingId { Val = numberingId }));
                target.AddBookmark("OverrideTarget");
                document.AddParagraph("Override reference: ")._paragraph.Append(BuildSimpleField(" REF OverrideTarget \\n ", "stale-n"));
                document.AddParagraph("Override full reference: ")._paragraph.Append(BuildSimpleField(" REF OverrideTarget \\w ", "stale-w"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(2, report.TotalCount);
                Assert.Equal(2, report.UpdatedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(new[] { "I", "Article I" }, report.Results.Select(result => result.ResultText).ToArray());
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_PreservesReferenceListClosingParentheses() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.ReferenceListNumbers.Parentheses.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                int numberingId = AddReferenceNumberingDefinition(document, "(%1)");
                WordParagraph target = document.AddParagraph("Parenthesized target");
                target._paragraph.ParagraphProperties = new ParagraphProperties(
                    new NumberingProperties(
                        new NumberingLevelReference { Val = 0 },
                        new NumberingId { Val = numberingId }));
                target.AddBookmark("ParenthesizedTarget");
                document.AddParagraph("Full reference: ")._paragraph.Append(BuildSimpleField(" REF ParenthesizedTarget \\w ", "stale-w"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                WordFieldUpdateResult result = Assert.Single(report.Results);
                Assert.Equal(WordFieldType.Ref, result.FieldType);
                Assert.Equal(WordFieldUpdateStatus.Updated, result.Status);
                Assert.Equal("(1)", result.ResultText);
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesRefListNumberSwitchesFromParagraphStyles() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.ReferenceListNumbers.ParagraphStyles.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                int numberingId = AddReferenceNumberingDefinition(document, "%1", "Article %1.%2");

                AddReferenceNumberingStyle(document, "StyledNumberedLevel0", null, numberingId, 0);
                AddReferenceNumberingStyle(document, "StyledNumberedLevel1", "StyledNumberedLevel0", null, 1);

                AddStyledParagraph(document, "Styled parent one", "StyledNumberedLevel0");
                AddStyledParagraph(document, "Styled parent two", "StyledNumberedLevel0");
                AddStyledParagraph(document, "Styled child target", "StyledNumberedLevel1").AddBookmark("StyledChild");

                WordParagraph source = AddStyledParagraph(document, "Styled child source relative: ", "StyledNumberedLevel1");
                source._paragraph.Append(BuildSimpleField(" REF StyledChild \\r ", "stale-r"));

                document.AddParagraph("Style no context: ")._paragraph.Append(BuildSimpleField(" REF StyledChild \\n ", "stale-n"));
                document.AddParagraph("Style full context: ")._paragraph.Append(BuildSimpleField(" REF StyledChild \\w ", "stale-w"));
                document.AddParagraph("Style full context without text: ")._paragraph.Append(BuildSimpleField(" REF StyledChild \\w \\t ", "stale-wt"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(4, report.TotalCount);
                Assert.Equal(4, report.UpdatedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);
                Assert.Equal(
                    new[] { "1", "1", "Article 2.1", "2.1" },
                    report.Results.Select(result => result.ResultText).ToArray());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());

                var fields = document.InspectFields()
                    .Where(field => field.FieldType == WordFieldType.Ref)
                    .Select(field => field.ResultText)
                    .ToArray();

                Assert.Equal(new[] { "1", "1", "Article 2.1", "2.1" }, fields);
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesRefListNumberSwitchesFromNumberingLevelParagraphStyles() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.ReferenceListNumbers.NumberingLevelParagraphStyles.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                AddReferenceParagraphStyle(document, "LinkedNumberedLevel0");
                AddReferenceParagraphStyle(document, "LinkedNumberedLevel1");
                AddReferenceNumberingDefinitionWithLevelStyles(
                    document,
                    new[] { "%1", "Clause %1.%2" },
                    new[] { "LinkedNumberedLevel0", "LinkedNumberedLevel1" });

                AddStyledParagraph(document, "Linked parent one", "LinkedNumberedLevel0");
                AddStyledParagraph(document, "Linked parent two", "LinkedNumberedLevel0");
                AddStyledParagraph(document, "Linked child target", "LinkedNumberedLevel1").AddBookmark("LinkedChild");

                WordParagraph source = AddStyledParagraph(document, "Linked child source relative: ", "LinkedNumberedLevel1");
                source._paragraph.Append(BuildSimpleField(" REF LinkedChild \\r ", "stale-r"));

                document.AddParagraph("Linked style no context: ")._paragraph.Append(BuildSimpleField(" REF LinkedChild \\n ", "stale-n"));
                document.AddParagraph("Linked style full context: ")._paragraph.Append(BuildSimpleField(" REF LinkedChild \\w ", "stale-w"));
                document.AddParagraph("Linked style full context without text: ")._paragraph.Append(BuildSimpleField(" REF LinkedChild \\w \\t ", "stale-wt"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(4, report.TotalCount);
                Assert.Equal(4, report.UpdatedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);
                Assert.Equal(
                    new[] { "1", "1", "Clause 2.1", "2.1" },
                    report.Results.Select(result => result.ResultText).ToArray());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());

                var fields = document.InspectFields()
                    .Where(field => field.FieldType == WordFieldType.Ref)
                    .Select(field => field.ResultText)
                    .ToArray();

                Assert.Equal(new[] { "1", "1", "Clause 2.1", "2.1" }, fields);
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesSequenceFieldsForGeneratedCaptionsAndReferences() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.CaptionSequences.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                AddCaptionParagraph(document, "FigureNetwork", "Figure ", "Network diagram");
                document.AddParagraph("Figure ").AddField(WordFieldType.Seq, parameters: new List<string> { "Figure", "\\n" });
                document.AddParagraph("Reset figure ").AddField(WordFieldType.Seq, parameters: new List<string> { "Figure", "\\r 10" });
                document.AddParagraph("Current figure ").AddField(WordFieldType.Seq, parameters: new List<string> { "Figure", "\\c" });
                document.AddParagraph("Ordinal figure ").AddField(WordFieldType.Seq, WordFieldFormat.Ordinal, parameters: new List<string> { "Figure" });
                document.AddParagraph("Hex figure ").AddField(WordFieldType.Seq, WordFieldFormat.Hex, parameters: new List<string> { "Figure" });
                document.AddParagraph("Cardinal text figure ").AddField(WordFieldType.Seq, WordFieldFormat.CardText, parameters: new List<string> { "Figure" });
                document.AddParagraph("Ordinal text figure ").AddField(WordFieldType.Seq, WordFieldFormat.OrdText, parameters: new List<string> { "Figure" });
                document.AddParagraph("Dollar text figure ").AddField(WordFieldType.Seq, WordFieldFormat.DollarText, parameters: new List<string> { "Figure" });
                document.AddParagraph("See ").AddField(WordFieldType.Ref, parameters: new List<string> { "FigureNetwork", "\\h" });
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(10, report.TotalCount);
                Assert.Equal(10, report.UpdatedCount);
                Assert.Equal(0, report.SkippedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);

                Assert.Equal(new[] { "1", "2", "10", "10", "11th", "C", "thirteen", "fourteenth", "fifteen and 00/100" }, report.Results
                    .Where(result => result.FieldType == WordFieldType.Seq)
                    .Select(result => result.ResultText)
                    .ToArray());

                Assert.Contains(report.Results, result =>
                    result.FieldType == WordFieldType.Ref &&
                    result.Status == WordFieldUpdateStatus.Updated &&
                    result.ResultText == "Figure 1Network diagram");

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields();

                Assert.Equal(new[] { "1", "2", "10", "10", "11th", "C", "thirteen", "fourteenth", "fifteen and 00/100" }, fields
                    .Where(field => field.FieldType == WordFieldType.Seq)
                    .Select(field => field.ResultText)
                    .ToArray());
                Assert.Contains(fields, field =>
                    field.FieldType == WordFieldType.Ref &&
                    field.ResultText == "Figure 1Network diagram");
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_KeepsHiddenSequenceFieldsHidden() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.HiddenSequence.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Hidden figure ")._paragraph.Append(BuildSimpleField(" SEQ Figure \\h ", "stale-hidden"));
                document.AddParagraph("Visible figure ")._paragraph.Append(BuildSimpleField(" SEQ Figure ", "stale-visible"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                WordFieldUpdateResult hidden = Assert.Single(report.Results, result =>
                    result.FieldType == WordFieldType.Seq &&
                    result.InstructionText.Contains("\\h", StringComparison.Ordinal));
                Assert.Equal(WordFieldUpdateStatus.Updated, hidden.Status);
                Assert.Equal(string.Empty, hidden.ResultText);
                WordFieldUpdateResult visible = Assert.Single(report.Results, result =>
                    result.FieldType == WordFieldType.Seq &&
                    !result.InstructionText.Contains("\\h", StringComparison.Ordinal));
                Assert.Equal("2", visible.ResultText);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, new WordLoadOptions { AccessMode = OfficeIMO.Drawing.DocumentAccessMode.ReadOnly })) {
                SimpleField hiddenField = document._wordprocessingDocument.MainDocumentPart!.Document!.Body!.Descendants<SimpleField>()
                    .Single(field => (field.Instruction?.Value ?? string.Empty).Contains("\\h", StringComparison.Ordinal));
                Assert.Equal(string.Empty, Assert.Single(hiddenField.Descendants<Text>()).Text);
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_IgnoresCharFormatWhenFormattingSequences() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.SequenceCharFormat.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Roman figure ")._paragraph.Append(BuildSimpleField(" SEQ Figure \\* Roman \\* CHARFORMAT ", "stale"));
                document.AddParagraph("Plain figure ")._paragraph.Append(BuildSimpleField(" SEQ Figure ", "stale"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(new[] { "I", "2" }, report.Results
                    .Where(result => result.FieldType == WordFieldType.Seq)
                    .Select(result => result.ResultText)
                    .ToArray());
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_ResetsSequenceFieldsAtHeadingLevel() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.SequenceHeadingReset.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Chapter A").SetStyle(WordParagraphStyles.Heading1);
                document.AddParagraph("Figure ").AddField(WordFieldType.Seq, parameters: new List<string> { "Figure", "\\s 1" });
                document.AddParagraph("Figure ").AddField(WordFieldType.Seq, parameters: new List<string> { "Figure", "\\s 1" });
                document.AddParagraph("Chapter B").SetStyle(WordParagraphStyles.Heading1);
                document.AddParagraph("Figure ").AddField(WordFieldType.Seq, parameters: new List<string> { "Figure", "\\s 1" });
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(3, report.TotalCount);
                Assert.Equal(3, report.UpdatedCount);
                Assert.Equal(new[] { "1", "2", "1" }, report.Results
                    .Where(result => result.FieldType == WordFieldType.Seq)
                    .Select(result => result.ResultText)
                    .ToArray());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Equal(new[] { "1", "2", "1" }, document.InspectFields()
                    .Where(field => field.FieldType == WordFieldType.Seq)
                    .Select(field => field.ResultText)
                    .ToArray());
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_ReportsUnsupportedSequenceSwitches() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.SequenceUnsupported.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Section figure ").AddField(WordFieldType.Seq, parameters: new List<string> { "Figure", "\\s 10" });
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                WordFieldUpdateResult result = Assert.Single(report.Results);
                Assert.Equal(WordFieldType.Seq, result.FieldType);
                Assert.Equal(WordFieldUpdateStatus.Unsupported, result.Status);
                Assert.Contains("\\s 10", result.Message, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesBoundedFormulaFields() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.Formulas.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Formula: ")._paragraph.Append(BuildSimpleField(" = 2 + 3 * (4 + 1) ", "stale"));
                document.AddParagraph("Formula with trailing switch: ")._paragraph.Append(BuildSimpleField(" = 2 + 3 \\* MERGEFORMAT ", "stale-mergeformat"));
                document.AddParagraph("Formula with Roman switch: ")._paragraph.Append(BuildSimpleField(" = 12 \\* Roman ", "stale-roman"));
                document.AddParagraph("Decimal formula: ")._paragraph.Append(BuildSimpleField(" = 7 / 2 ", "stale"));
                document.AddParagraph("Sum formula: ")._paragraph.Append(BuildSimpleField(" = SUM(1, 2, 3 + 4) ", "stale"));
                document.AddParagraph("Average formula: ")._paragraph.Append(BuildSimpleField(" = AVERAGE(2, 4, 6) ", "stale"));
                document.AddParagraph("Min formula: ")._paragraph.Append(BuildSimpleField(" = MIN(5, 3, 7) ", "stale"));
                document.AddParagraph("Max formula: ")._paragraph.Append(BuildSimpleField(" = MAX(5, 3, 7) ", "stale"));
                document.AddParagraph("Product formula: ")._paragraph.Append(BuildSimpleField(" = PRODUCT(2, 3, 4) ", "stale"));
                document.AddParagraph("Count formula: ")._paragraph.Append(BuildSimpleField(" = COUNT(2, 3, 4) ", "stale"));
                document.AddParagraph("Grouped percent formula: ")._paragraph.Append(BuildSimpleField(" = (20 + 5)% ", "stale"));
                document.AddParagraph("Function percent formula: ")._paragraph.Append(BuildSimpleField(" = ABS(-25)% ", "stale"));
                document.AddParagraph("Round/abs formula: ")._paragraph.Append(BuildSimpleField(" = ROUND(ABS(-2.345), 2) ", "stale"));
                document.AddParagraph("Negative-place round formula: ")._paragraph.Append(BuildSimpleField(" = ROUND(1265, -2) ", "stale"));
                document.AddParagraph("Positive integer formula: ")._paragraph.Append(BuildSimpleField(" = INT(7 / 2) ", "stale"));
                document.AddParagraph("Negative integer formula: ")._paragraph.Append(BuildSimpleField(" = INT(-3.2) ", "stale"));
                document.AddParagraph("Defined expression formula: ")._paragraph.Append(BuildSimpleField(" = DEFINED(1 + 2) ", "stale"));
                document.AddParagraph("Undefined expression formula: ")._paragraph.Append(BuildSimpleField(" = DEFINED(MEDIAN(1, 2, 3)) ", "stale"));
                document.AddParagraph("Unsupported formula: ")._paragraph.Append(BuildSimpleField(" = MEDIAN(1, 2, 3) ", "stale"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(19, report.TotalCount);
                Assert.Equal(18, report.UpdatedCount);
                Assert.Equal(1, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);

                Assert.Equal(new[] { "17", "5", "XII", "3.5", "10", "4", "3", "7", "24", "3", "0.25", "0.25", "2.35", "1300", "3", "-4", "1", "0" }, report.Results
                    .Where(result => result.FieldType == WordFieldType.Formula && result.Status == WordFieldUpdateStatus.Updated)
                    .Select(result => result.ResultText)
                    .ToArray());

                WordFieldUpdateResult unsupported = Assert.Single(report.Results, result =>
                    result.FieldType == WordFieldType.Formula &&
                    result.Status == WordFieldUpdateStatus.Unsupported);
                Assert.Contains("MEDIAN", unsupported.Message, StringComparison.OrdinalIgnoreCase);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields()
                    .Where(field => field.FieldType == WordFieldType.Formula)
                    .ToArray();

                Assert.Equal(new[] { "17", "5", "XII", "3.5", "10", "4", "3", "7", "24", "3", "0.25", "0.25", "2.35", "1300", "3", "-4", "1", "0", "stale" }, fields.Select(field => field.ResultText).ToArray());
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_EvaluatesHugeIntegralExponentInLogarithmicSteps() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.FormulaExponent.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Formula: ")._paragraph.Append(BuildSimpleField(" = 1 ^ 2147483647 ", "stale"));
                document.Save();
            }

            using WordDocument loaded = WordDocument.Load(filePath);
            WordFieldUpdateResult result = Assert.Single(loaded.UpdateFieldsAndGetReport().Results);

            Assert.Equal(WordFieldUpdateStatus.Updated, result.Status);
            Assert.Equal("1", result.ResultText);
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_EvaluatesLongExponentChainsWithoutParserRecursion() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.FormulaExponentChain.docx");
            string expression = string.Join(" ^ ", Enumerable.Repeat("1", 4096));
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Formula: ")._paragraph.Append(BuildSimpleField(" = " + expression + " ", "stale"));
                document.Save();
            }

            using WordDocument loaded = WordDocument.Load(filePath);
            WordFieldUpdateResult result = Assert.Single(loaded.UpdateFieldsAndGetReport().Results);

            Assert.Equal(WordFieldUpdateStatus.Updated, result.Status);
            Assert.Equal("1", result.ResultText);
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_AppliesUnarySignsAfterExponentiation() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.FormulaUnaryExponent.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Signed base: ")._paragraph.Append(BuildSimpleField(" = -2 ^ 2 ", "stale"));
                document.AddParagraph("Grouped base: ")._paragraph.Append(BuildSimpleField(" = (-2) ^ 2 ", "stale"));
                document.AddParagraph("Signed exponent: ")._paragraph.Append(BuildSimpleField(" = 2 ^ -2 ^ 2 ", "stale"));
                document.Save();
            }

            using WordDocument loaded = WordDocument.Load(filePath);
            WordFieldUpdateReport report = loaded.UpdateFieldsAndGetReport();

            Assert.Equal(new[] { "-4", "4", "0.0625" }, report.Results.Select(result => result.ResultText).ToArray());
            Assert.All(report.Results, result => Assert.Equal(WordFieldUpdateStatus.Updated, result.Status));
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_AppliesFormulaNumericPictureSwitches() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.FormulaNumericPictures.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Fixed decimal: ")._paragraph.Append(BuildSimpleField(" = 7 / 2 \\# \"0.00\" ", "stale-decimal"));
                document.AddParagraph("Trailing mergeformat: ")._paragraph.Append(BuildSimpleField(" = 7 / 2 \\# \"0.00\" \\* MERGEFORMAT ", "stale-trailing-mergeformat"));
                document.AddParagraph("Thousands: ")._paragraph.Append(BuildSimpleField(" = 1234.5 \\# \"#,##0.00\" ", "stale-thousands"));
                document.AddParagraph("Percent: ")._paragraph.Append(BuildSimpleField(" = 1 / 4 \\# \"0.0%\" ", "stale-percent"));
                document.AddParagraph("Literal suffix: ")._paragraph.Append(BuildSimpleField(" = 5 \\# \"0.00 USD\" ", "stale-literal"));
                document.AddParagraph("Color tag: ")._paragraph.Append(BuildSimpleField(" = 5 \\# \"[Red]0.00\" ", "stale-color"));
                document.AddParagraph("Negative section: ")._paragraph.Append(BuildSimpleField(" = -5 \\# \"0.00;[Red](0.00);0.00\" ", "stale-negative"));
                document.AddParagraph("Zero section: ")._paragraph.Append(BuildSimpleField(" = 0 \\# \"0.00;[Red](0.00);Zero\" ", "stale-zero"));
                document.AddParagraph("Conditional high section: ")._paragraph.Append(BuildSimpleField(" = 150 \\# \"[>=100]0.0 high;[<100]0.0 low;0.0\" ", "stale-high"));
                document.AddParagraph("Conditional low section: ")._paragraph.Append(BuildSimpleField(" = 50 \\# \"[>=100]0.0 high;[<100]0.0 low;0.0\" ", "stale-low"));
                document.AddParagraph("Escaped literal suffix: ")._paragraph.Append(BuildSimpleField(" = 5 \\# \"0.00\\ USD\" ", "stale-escape"));
                document.AddParagraph("Fill token: ")._paragraph.Append(BuildSimpleField(" = 5 \\# \"0.00* \" ", "stale-fill"));
                document.AddParagraph("Unsupported dangling fill: ")._paragraph.Append(BuildSimpleField(" = 5 \\# \"0.00*\" ", "stale-dangling-fill"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(13, report.TotalCount);
                Assert.Equal(12, report.UpdatedCount);
                Assert.Equal(1, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);

                Assert.Equal(new[] { "3.50", "3.50", "1,234.50", "25.0%", "5.00 USD", "5.00", "(5.00)", "Zero", "150.0 high", "50.0 low", "5.00 USD", "5.00" }, report.Results
                    .Where(result => result.FieldType == WordFieldType.Formula && result.Status == WordFieldUpdateStatus.Updated)
                    .Select(result => result.ResultText)
                    .ToArray());

                WordFieldUpdateResult unsupported = Assert.Single(report.Results, result =>
                    result.FieldType == WordFieldType.Formula &&
                    result.Status == WordFieldUpdateStatus.Unsupported);
                Assert.Contains("numeric picture", unsupported.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("unsupported", unsupported.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("fill", unsupported.Message, StringComparison.OrdinalIgnoreCase);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields()
                    .Where(field => field.FieldType == WordFieldType.Formula)
                    .ToArray();

                Assert.Equal(new[] { "3.50", "3.50", "1,234.50", "25.0%", "5.00 USD", "5.00", "(5.00)", "Zero", "150.0 high", "50.0 low", "5.00 USD", "5.00", "stale-dangling-fill" }, fields.Select(field => field.ResultText).ToArray());
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesConditionalFormulaFields() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.ConditionalFormulas.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Greater-than formula: ")._paragraph.Append(BuildSimpleField(" = 2 > 1 ", "stale"));
                document.AddParagraph("Less-than-or-equal formula: ")._paragraph.Append(BuildSimpleField(" = 2 <= 1 ", "stale"));
                document.AddParagraph("Not-equal formula: ")._paragraph.Append(BuildSimpleField(" = 3 <> 4 ", "stale"));
                document.AddParagraph("Greater-than-or-equal formula: ")._paragraph.Append(BuildSimpleField(" = 3 >= 4 ", "stale"));
                document.AddParagraph("If formula: ")._paragraph.Append(BuildSimpleField(" = IF(2 > 1, 10, 20) ", "stale"));
                document.AddParagraph("If true skips unsupported false branch: ")._paragraph.Append(BuildSimpleField(" = IF(TRUE, 10, MEDIAN(1, 2, 3)) ", "stale"));
                document.AddParagraph("If false skips unsupported true branch: ")._paragraph.Append(BuildSimpleField(" = IF(FALSE, MEDIAN(1, 2, 3), 20) ", "stale"));
                document.AddParagraph("And/not formula: ")._paragraph.Append(BuildSimpleField(" = IF(AND(1 = 1, NOT(FALSE())), 7, 9) ", "stale"));
                document.AddParagraph("And skips unsupported after false: ")._paragraph.Append(BuildSimpleField(" = AND(FALSE, MEDIAN(1, 2, 3)) ", "stale"));
                document.AddParagraph("Or formula: ")._paragraph.Append(BuildSimpleField(" = IF(OR(FALSE, 3 < 2), 7, 9) ", "stale"));
                document.AddParagraph("Or skips unsupported after true: ")._paragraph.Append(BuildSimpleField(" = OR(TRUE, MEDIAN(1, 2, 3)) ", "stale"));
                document.AddParagraph("Modulo formula: ")._paragraph.Append(BuildSimpleField(" = MOD(17, 5) ", "stale"));
                document.AddParagraph("Sign formula: ")._paragraph.Append(BuildSimpleField(" = SIGN(-4) ", "stale"));
                document.AddParagraph("Invalid true formula: ")._paragraph.Append(BuildSimpleField(" = TRUE(1) ", "stale"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(14, report.TotalCount);
                Assert.Equal(13, report.UpdatedCount);
                Assert.Equal(1, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);

                Assert.Equal(new[] { "1", "0", "1", "0", "10", "10", "20", "7", "0", "9", "1", "2", "-1" }, report.Results
                    .Where(result => result.FieldType == WordFieldType.Formula && result.Status == WordFieldUpdateStatus.Updated)
                    .Select(result => result.ResultText)
                    .ToArray());

                WordFieldUpdateResult unsupported = Assert.Single(report.Results, result =>
                    result.FieldType == WordFieldType.Formula &&
                    result.Status == WordFieldUpdateStatus.Unsupported);
                Assert.Contains("TRUE", unsupported.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Contains("0", unsupported.Message, StringComparison.Ordinal);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields()
                    .Where(field => field.FieldType == WordFieldType.Formula)
                    .ToArray();

                Assert.Equal(new[] { "1", "0", "1", "0", "10", "10", "20", "7", "0", "9", "1", "2", "-1", "stale" }, fields.Select(field => field.ResultText).ToArray());
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesFormulaFieldsWithSemicolonArgumentSeparators() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.FormulaSemicolonSeparators.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Semicolon sum: ")._paragraph.Append(BuildSimpleField(" = SUM(1; 2; 3 + 4) ", "stale-sum"));
                document.AddParagraph("Semicolon if: ")._paragraph.Append(BuildSimpleField(" = IF(AND(1 = 1; NOT(FALSE())); 10; 20) ", "stale-if"));
                document.AddParagraph("Semicolon modulo: ")._paragraph.Append(BuildSimpleField(" = MOD(17; 5) ", "stale-mod"));

                WordTable table = document.AddTable(2, 3);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "10";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "20";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "30";
                table.Rows[1].Cells[1].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = SUM(A1; B1; R2C1) ", "stale-table"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(4, report.TotalCount);
                Assert.Equal(4, report.UpdatedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);
                Assert.Equal(new[] { "10", "10", "2", "60" }, report.Results.Select(result => result.ResultText).ToArray());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());

                var fields = document.InspectFields()
                    .Where(field => field.FieldType == WordFieldType.Formula)
                    .ToArray();

                Assert.Equal(new[] { "10", "10", "2", "60" }, fields.Select(field => field.ResultText).ToArray());
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesFormulaFieldsWithPercentLiterals() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.FormulaPercentLiterals.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Percent literal: ")._paragraph.Append(BuildSimpleField(" = 25% ", "stale-percent"));
                document.AddParagraph("Percent multiplication: ")._paragraph.Append(BuildSimpleField(" = 100 * 25% ", "stale-multiply"));
                document.AddParagraph("Percent conditional: ")._paragraph.Append(BuildSimpleField(" = IF(50% = 0.5, 1, 0) ", "stale-if"));
                document.AddParagraph("Percent picture: ")._paragraph.Append(BuildSimpleField(" = 25% \\# \"0%\" ", "stale-picture"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(4, report.TotalCount);
                Assert.Equal(4, report.UpdatedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);
                Assert.Equal(new[] { "0.25", "25", "1", "25%" }, report.Results.Select(result => result.ResultText).ToArray());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());

                var fields = document.InspectFields()
                    .Where(field => field.FieldType == WordFieldType.Formula)
                    .ToArray();

                Assert.Equal(new[] { "0.25", "25", "1", "25%" }, fields.Select(field => field.ResultText).ToArray());
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_ReadsPercentValuedTableCellsForFormulaReferences() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.TableFormulas.PercentCells.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(3, 3);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "25%";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "50 %";
                table.Rows[0].Cells[2].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = SUM(LEFT) \\# \"0%\" ", "stale-left"));
                table.Rows[1].Cells[0].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = SUM(A1:B1) \\# \"0%\" ", "stale-range"));
                table.Rows[1].Cells[1].Paragraphs[0].Text = "12.5%";
                table.Rows[1].Cells[2].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = PRODUCT(A1:B1) \\# \"0.0%\" ", "stale-product"));
                table.Rows[2].Cells[0].Paragraphs[0].Text = "20%";
                table.Rows[2].Cells[1].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = SUM(ABOVE) \\# \"0.0%\" ", "stale-above"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(4, report.TotalCount);
                Assert.Equal(4, report.UpdatedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);
                Assert.Equal(new[] { "75%", "75%", "12.5%", "62.5%" }, report.Results.Select(result => result.ResultText).ToArray());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());

                var fields = document.InspectFields()
                    .Where(field => field.FieldType == WordFieldType.Formula)
                    .ToArray();

                Assert.Equal(new[] { "75%", "75%", "12.5%", "62.5%" }, fields.Select(field => field.ResultText).ToArray());
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesImportedStyleComplexFormulaFieldsWithPercentTableCells() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.TableFormulas.ImportedComplexPercentCells.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(3, 3);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "25%";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "50 %";
                AddComplexField(table.Rows[0].Cells[2].Paragraphs[0]._paragraph, "stale-left", " = SUM(", "LEFT", ") \\# \"0%\" ");
                AddComplexField(table.Rows[1].Cells[0].Paragraphs[0]._paragraph, "stale-range", " = SUM(", "A1:B1", ") \\# \"0%\" ");
                table.Rows[1].Cells[1].Paragraphs[0].Text = "12.5%";
                AddComplexField(table.Rows[1].Cells[2].Paragraphs[0]._paragraph, "stale-product", " = PRODUCT(", "A1:B1", ") \\# \"0.0%\" ");
                table.Rows[2].Cells[0].Paragraphs[0].Text = "20%";
                AddComplexField(table.Rows[2].Cells[1].Paragraphs[0]._paragraph, "stale-above", " = SUM(", "ABOVE", ") \\# \"0.0%\" ");
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(4, report.TotalCount);
                Assert.Equal(4, report.UpdatedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);
                Assert.Equal(new[] { "75%", "75%", "12.5%", "62.5%" }, report.Results.Select(result => result.ResultText).ToArray());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());

                var fields = document.InspectFields()
                    .Where(field => field.FieldType == WordFieldType.Formula)
                    .ToArray();

                Assert.All(fields, field => Assert.Equal(WordFieldRepresentation.Complex, field.Representation));
                Assert.Equal(new[] { "75%", "75%", "12.5%", "62.5%" }, fields.Select(field => field.ResultText).ToArray());
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UpdatesTableRelativeFormulaFields() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.TableFormulas.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(4, 5);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "10";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "2";
                table.Rows[0].Cells[2].Paragraphs[0].Text = "4";
                table.Rows[0].Cells[3].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = SUM(BELOW) ", "stale"));
                table.Rows[0].Cells[4].Paragraphs[0].Text = "7";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "20";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "4";
                table.Rows[1].Cells[2].Paragraphs[0].Text = "6";
                table.Rows[1].Cells[3].Paragraphs[0].Text = "1";
                table.Rows[1].Cells[4].Paragraphs[0].Text = "9";
                table.Rows[2].Cells[0].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = SUM(ABOVE) ", "stale"));
                table.Rows[2].Cells[1].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = AVERAGE(ABOVE) ", "stale"));
                table.Rows[2].Cells[2].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = MIN(ABOVE) ", "stale"));
                table.Rows[2].Cells[3].Paragraphs[0].Text = "2";
                table.Rows[2].Cells[4].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = COUNT(ABOVE) ", "stale"));
                table.Rows[3].Cells[0].Paragraphs[0].Text = "3";
                table.Rows[3].Cells[1].Paragraphs[0].Text = "5";
                table.Rows[3].Cells[2].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = SUM(LEFT) ", "stale"));
                table.Rows[3].Cells[3].Paragraphs[0].Text = "3";
                table.Rows[3].Cells[4].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = COUNT(LEFT) ", "stale"));
                document.AddParagraph("Outside table: ")._paragraph.Append(BuildSimpleField(" = SUM(ABOVE) ", "stale-outside"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(8, report.TotalCount);
                Assert.Equal(7, report.UpdatedCount);
                Assert.Equal(1, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);

                Assert.Equal(new[] { "6", "30", "3", "4", "2", "8", "4" }, report.Results
                    .Where(result => result.FieldType == WordFieldType.Formula && result.Status == WordFieldUpdateStatus.Updated)
                    .Select(result => result.ResultText)
                    .ToArray());

                WordFieldUpdateResult unsupported = Assert.Single(report.Results, result =>
                    result.FieldType == WordFieldType.Formula &&
                    result.Status == WordFieldUpdateStatus.Unsupported);
                Assert.Contains("inside a table cell", unsupported.Message, StringComparison.OrdinalIgnoreCase);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                var fields = document.InspectFields()
                    .Where(field => field.FieldType == WordFieldType.Formula)
                    .ToArray();

                Assert.Equal(new[] { "6", "30", "3", "4", "2", "8", "4", "stale-outside" }, fields.Select(field => field.ResultText).ToArray());
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UsesVisualColumnsForMergedTableFormulaReferences() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.TableFormulas.MergedColumns.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(2, 3);
                WordTableCell spannedCell = table.Rows[0].Cells[0];
                WordTableCell omittedCell = table.Rows[0].Cells[1];
                WordTableCell thirdColumnCell = table.Rows[0].Cells[2];

                spannedCell.Paragraphs[0].Text = "10";
                TableCellProperties properties = spannedCell._tableCell.GetFirstChild<TableCellProperties>()
                    ?? spannedCell._tableCell.PrependChild(new TableCellProperties());
                properties.GridSpan = new GridSpan { Val = 2 };
                thirdColumnCell.Paragraphs[0].Text = "5";
                omittedCell._tableCell.Remove();

                table.Rows[1].Cells[0].Paragraphs[0].Text = "1";
                table.Rows[1].Cells[1].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = SUM(ABOVE) ", "stale-second-column"));
                table.Rows[1].Cells[2].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = SUM(ABOVE) ", "stale-third-column"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(2, report.TotalCount);
                Assert.Equal(2, report.UpdatedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);
                Assert.Equal(new[] { "10", "5" }, report.Results.Select(result => result.ResultText).ToArray());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());

                var fields = document.InspectFields()
                    .Where(field => field.FieldType == WordFieldType.Formula)
                    .ToArray();

                Assert.Equal(new[] { "10", "5" }, fields.Select(field => field.ResultText).ToArray());
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_SkipsVerticalMergeContinuationsForTableFormulaReferences() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.TableFormulas.VerticalMerge.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(4, 2);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "10";
                table.Rows[0].Cells[1].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = SUM(BELOW) ", "stale-below"));
                table.Rows[1].Cells[0].Paragraphs[0].Text = "999";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "4";
                table.Rows[2].Cells[1].Paragraphs[0].Text = "777";
                table.Rows[3].Cells[0].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = SUM(ABOVE) ", "stale-above"));
                table.Rows[3].Cells[1].Paragraphs[0].Text = "6";

                SetVerticalMerge(table.Rows[0].Cells[0], MergedCellValues.Restart);
                SetVerticalMerge(table.Rows[1].Cells[0], MergedCellValues.Continue);
                SetVerticalMerge(table.Rows[1].Cells[1], MergedCellValues.Restart);
                SetVerticalMerge(table.Rows[2].Cells[1], null);

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(2, report.TotalCount);
                Assert.Equal(2, report.UpdatedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);
                Assert.Equal(new[] { "10", "10" }, report.Results.Select(result => result.ResultText).ToArray());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());

                var fields = document.InspectFields()
                    .Where(field => field.FieldType == WordFieldType.Formula)
                    .ToArray();

                Assert.Equal(new[] { "10", "10" }, fields.Select(field => field.ResultText).ToArray());
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UsesGridBeforeOffsetsForTableFormulaReferences() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.TableFormulas.GridBefore.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(2, 3);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "10";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "20";
                table.Rows[0].Cells[2].Paragraphs[0].Text = "30";
                table.Rows[1].Cells[0].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = SUM(ABOVE) ", "stale-middle"));
                table.Rows[1].Cells[1].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = SUM(ABOVE) ", "stale-last"));
                table.Rows[1].Cells[2]._tableCell.Remove();
                SetGridBefore(table.Rows[1], 1);
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(2, report.TotalCount);
                Assert.Equal(2, report.UpdatedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);
                Assert.Equal(new[] { "20", "30" }, report.Results.Select(result => result.ResultText).ToArray());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());

                var fields = document.InspectFields()
                    .Where(field => field.FieldType == WordFieldType.Formula)
                    .ToArray();

                Assert.Equal(new[] { "20", "30" }, fields.Select(field => field.ResultText).ToArray());
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UsesExplicitTableCellReferencesAndRanges() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.TableFormulas.CellReferences.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(3, 3);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "10";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "20";
                table.Rows[0].Cells[2].Paragraphs[0].Text = "30";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "40";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "50";
                table.Rows[1].Cells[2]._tableCell.Remove();
                SetGridBefore(table.Rows[1], 1);
                table.Rows[2].Cells[0].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = A1 + B1 ", "stale-single"));
                table.Rows[2].Cells[1].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = SUM(B1:C2) ", "stale-range"));
                table.Rows[2].Cells[2].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = C2 * 2 ", "stale-shifted"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(3, report.TotalCount);
                Assert.Equal(3, report.UpdatedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);
                Assert.Equal(new[] { "30", "140", "100" }, report.Results.Select(result => result.ResultText).ToArray());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());

                var fields = document.InspectFields()
                    .Where(field => field.FieldType == WordFieldType.Formula)
                    .ToArray();

                Assert.Equal(new[] { "30", "140", "100" }, fields.Select(field => field.ResultText).ToArray());
            }
        }

        [Fact]
        public void Test_UpdateFieldsAndGetReport_UsesRnCnTableCellReferencesAndRanges() {
            string filePath = Path.Combine(_directoryWithFiles, "FieldUpdate.TableFormulas.RnCnCellReferences.docx");

            using (WordDocument document = WordDocument.Create(filePath)) {
                WordTable table = document.AddTable(3, 3);
                table.Rows[0].Cells[0].Paragraphs[0].Text = "10";
                table.Rows[0].Cells[1].Paragraphs[0].Text = "20";
                table.Rows[0].Cells[2].Paragraphs[0].Text = "30";
                table.Rows[1].Cells[0].Paragraphs[0].Text = "40";
                table.Rows[1].Cells[1].Paragraphs[0].Text = "50";
                table.Rows[1].Cells[2].Paragraphs[0].Text = "60";
                table.Rows[2].Cells[0].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = R1C1 + R2C2 ", "stale-rncn-single"));
                table.Rows[2].Cells[1].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = SUM(R1C1:R2C2) ", "stale-rncn-range"));
                table.Rows[2].Cells[2].Paragraphs[0]._paragraph.Append(BuildSimpleField(" = R2C3 * 2 ", "stale-rncn-third"));
                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                WordFieldUpdateReport report = document.UpdateFieldsAndGetReport();

                Assert.Equal(3, report.TotalCount);
                Assert.Equal(3, report.UpdatedCount);
                Assert.Equal(0, report.UnsupportedCount);
                Assert.Equal(0, report.ParseErrorCount);
                Assert.Equal(new[] { "60", "120", "120" }, report.Results.Select(result => result.ResultText).ToArray());

                document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                Assert.Empty(document.ValidateDocument());

                var fields = document.InspectFields()
                    .Where(field => field.FieldType == WordFieldType.Formula)
                    .ToArray();

                Assert.Equal(new[] { "60", "120", "120" }, fields.Select(field => field.ResultText).ToArray());
            }
        }

        private static SimpleField BuildSimpleField(string instruction, string resultText) {
            return new SimpleField(
                new Run(
                    new Text(resultText) { Space = SpaceProcessingModeValues.Preserve })) {
                Instruction = instruction
            };
        }

        private static void AddComplexField(Paragraph paragraph, string resultText, params string[] instructionParts) {
            paragraph.Append(new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }));
            foreach (string instructionPart in instructionParts) {
                paragraph.Append(new Run(new FieldCode { Text = instructionPart, Space = SpaceProcessingModeValues.Preserve }));
            }

            paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text(resultText) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }

        private static void AppendBookmarkedParagraph(OpenXmlCompositeElement root, string bookmarkName, string bookmarkId, string text) {
            root.Append(new Paragraph(
                new BookmarkStart { Name = bookmarkName, Id = bookmarkId },
                new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }),
                new BookmarkEnd { Id = bookmarkId }));
        }

        private static void AddNestedComplexFields(
            Paragraph paragraph,
            string outerInstruction,
            string innerInstruction,
            string outerPrefix,
            string innerResult,
            string outerSuffix) {
            paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = outerInstruction, Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text(outerPrefix) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = innerInstruction, Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text(innerResult) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                new Run(new Text(outerSuffix) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }

        private static void AddNestedInstructionComplexField(
            Paragraph paragraph,
            string outerInstructionPrefix,
            string innerInstruction,
            string innerResult,
            string outerInstructionSuffix,
            string outerResult) {
            paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = outerInstructionPrefix, Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
                new Run(new FieldCode { Text = innerInstruction, Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text(innerResult) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }),
                new Run(new FieldCode { Text = outerInstructionSuffix, Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text(outerResult) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }

        private static void SetVerticalMerge(WordTableCell cell, MergedCellValues? value) {
            TableCellProperties properties = cell._tableCell.GetFirstChild<TableCellProperties>()
                ?? cell._tableCell.PrependChild(new TableCellProperties());
            properties.VerticalMerge = new VerticalMerge { Val = value };
        }

        private static void SetGridBefore(WordTableRow row, int value) {
            row._tableRow.TableRowProperties ??= new TableRowProperties();
            row._tableRow.TableRowProperties.Append(new GridBefore { Val = value });
        }

        private static void AddReferenceNumberingLevel(WordList list, string levelText) {
            WordListLevel level = new WordListLevel(WordListLevelKind.DecimalDot) {
                LevelText = levelText
            };
            list.Numbering.AddLevel(level);
        }

        private static int AddReferenceNumberingDefinition(WordDocument document, params string[] levelTexts) {
            return AddReferenceNumberingDefinition(document, levelTexts, null);
        }

        private static int AddReferenceNumberingDefinitionWithLevelStyles(WordDocument document, string[] levelTexts, string?[] paragraphStyleIds) {
            return AddReferenceNumberingDefinition(document, levelTexts, paragraphStyleIds);
        }

        private static int AddReferenceNumberingDefinition(WordDocument document, IReadOnlyList<string> levelTexts, IReadOnlyList<string?>? paragraphStyleIds) {
            MainDocumentPart mainPart = document._wordprocessingDocument.MainDocumentPart!;
            NumberingDefinitionsPart numberingPart = mainPart.NumberingDefinitionsPart ?? mainPart.AddNewPart<NumberingDefinitionsPart>();
            numberingPart.Numbering ??= new Numbering();
            Numbering numbering = numberingPart.Numbering;
            int abstractId = numbering.Elements<AbstractNum>()
                .Select(abstractNum => abstractNum.AbstractNumberId?.Value ?? 0)
                .DefaultIfEmpty()
                .Max() + 1;
            int numberingId = numbering.Elements<NumberingInstance>()
                .Select(instance => instance.NumberID?.Value ?? 0)
                .DefaultIfEmpty()
                .Max() + 1;

            var abstractNum = new AbstractNum { AbstractNumberId = abstractId };
            for (int levelIndex = 0; levelIndex < levelTexts.Count; levelIndex++) {
                var level = new Level {
                    LevelIndex = levelIndex
                };
                level.Append(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Decimal });

                if (paragraphStyleIds != null &&
                    levelIndex < paragraphStyleIds.Count &&
                    !string.IsNullOrWhiteSpace(paragraphStyleIds[levelIndex])) {
                    level.Append(new ParagraphStyleIdInLevel { Val = paragraphStyleIds[levelIndex] });
                }

                level.Append(
                    new LevelText { Val = levelTexts[levelIndex] },
                    new LevelJustification { Val = LevelJustificationValues.Left });
                abstractNum.Append(level);
            }

            numbering.Append(abstractNum);
            numbering.Append(new NumberingInstance(
                new AbstractNumId { Val = abstractId }) {
                NumberID = numberingId
            });

            return numberingId;
        }

        private static void AddReferenceParagraphStyle(WordDocument document, string styleId) {
            var style = new Style(new StyleName { Val = styleId }) {
                Type = StyleValues.Paragraph,
                StyleId = styleId
            };

            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
        }

        private static void AddReferenceNumberingStyle(WordDocument document, string styleId, string? basedOnStyleId, int? numberId, int? level) {
            NumberingProperties numberingProperties = new NumberingProperties();
            if (level.HasValue) {
                numberingProperties.Append(new NumberingLevelReference { Val = level.Value });
            }

            if (numberId.HasValue) {
                numberingProperties.Append(new NumberingId { Val = numberId.Value });
            }

            Style style = new Style(
                new StyleName { Val = styleId },
                new StyleParagraphProperties(numberingProperties)) {
                Type = StyleValues.Paragraph,
                StyleId = styleId
            };

            if (!string.IsNullOrWhiteSpace(basedOnStyleId)) {
                style.InsertAfter(new BasedOn { Val = basedOnStyleId }, style.GetFirstChild<StyleName>());
            }

            document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Append(style);
        }

        private static WordParagraph AddStyledParagraph(WordDocument document, string text, string styleId) {
            WordParagraph paragraph = document.AddParagraph(text);
            paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
            paragraph._paragraph.ParagraphProperties.ParagraphStyleId = new ParagraphStyleId { Val = styleId };
            return paragraph;
        }

        private static void AddCaptionParagraph(WordDocument document, string bookmarkName, string label, string text) {
            string id = "9101";

            Paragraph paragraph = new Paragraph(
                new BookmarkStart { Name = bookmarkName, Id = id },
                new Run(new Text(label) { Space = SpaceProcessingModeValues.Preserve }),
                BuildSimpleField(" SEQ Figure ", "stale"),
                new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }),
                new BookmarkEnd { Id = id });

            document._document.Body!.Append(paragraph);
        }

        private static string ToRoman(int value) {
            if (value <= 0 || value > 3999) {
                return value.ToString(System.Globalization.CultureInfo.InvariantCulture);
            }

            (int Number, string Numeral)[] map = {
                (1000, "M"), (900, "CM"), (500, "D"), (400, "CD"),
                (100, "C"), (90, "XC"), (50, "L"), (40, "XL"),
                (10, "X"), (9, "IX"), (5, "V"), (4, "IV"), (1, "I")
            };

            var builder = new System.Text.StringBuilder();
            int remaining = value;
            foreach ((int number, string numeral) in map) {
                while (remaining >= number) {
                    builder.Append(numeral);
                    remaining -= number;
                }
            }

            return builder.ToString();
        }

        private static void AssertUpdated(WordFieldUpdateReport report, WordFieldType fieldType, string expectedResult) {
            WordFieldUpdateResult result = Assert.Single(report.Results, item => item.FieldType == fieldType);
            Assert.Equal(WordFieldUpdateStatus.Updated, result.Status);
            Assert.Equal(expectedResult, result.ResultText);
        }

        private static void AssertDocPropertyUpdated(WordFieldUpdateReport report, string propertyName, string expectedResult) {
            WordFieldUpdateResult result = Assert.Single(report.Results, item =>
                item.FieldType == WordFieldType.DocProperty &&
                item.InstructionText.Contains(propertyName, StringComparison.Ordinal));

            Assert.Equal(WordFieldUpdateStatus.Updated, result.Status);
            Assert.Equal(expectedResult, result.ResultText);
        }

        private static void AssertDocVariableUpdated(WordFieldUpdateReport report, string variableName, string expectedResult) {
            WordFieldUpdateResult result = Assert.Single(report.Results, item =>
                item.FieldType == WordFieldType.DocVariable &&
                item.InstructionText.Contains(variableName, StringComparison.Ordinal));

            Assert.Equal(WordFieldUpdateStatus.Updated, result.Status);
            Assert.Equal(expectedResult, result.ResultText);
        }
    }
}
