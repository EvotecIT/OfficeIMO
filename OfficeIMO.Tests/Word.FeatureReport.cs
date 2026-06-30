using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_WordFeatureReport_DetectsEditableAndAdvancedFeatures() {
            string filePath = Path.Combine(_directoryWithFiles, "WordFeatureReport.Advanced.docm");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Intro");
                document.AddTable(2, 2);
                document.AddParagraph("Footnote anchor").AddFootNote("Footnote body");
                document.AddParagraph("Endnote anchor").AddEndNote("Endnote body");
                document.AddParagraph("Comment target").AddComment("Jane Doe", "JD", "Review note");
                var revisionParagraph = document.AddParagraph();
                revisionParagraph.AddInsertedText("Inserted text", "Jane Doe");
                revisionParagraph.AddDeletedText("Deleted text", "Jane Doe");
                document.AddParagraph("Choice").AddCheckBox(true, "Approval", "ApprovalTag");
                document.AddParagraph("External ").AddHyperLink("site", new Uri("https://example.com"));
                document.AddEmbeddedFragment("<html><body><p>Imported</p></body></html>", WordAlternativeFormatImportPartType.Html);
                document.AddMacro(System.Text.Encoding.ASCII.GetBytes("OfficeIMO macro placeholder"));
                document.ApplicationProperties.DigitalSignature = new DigitalSignature();
                document.Save(false, new WordSaveOptions { SignedDocumentPolicy = WordSignedDocumentSavePolicy.AllowSignatureInvalidation });
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordFeatureReport report = document.InspectFeatures();

                Assert.Contains(report.EditableFeatures, feature => feature.Name == "Paragraphs" && feature.Count > 0);
                Assert.Contains(report.EditableFeatures, feature => feature.Name == "Tables" && feature.Count == 1);
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Footnotes" && feature.Count == 1);
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Endnotes" && feature.Count == 1);
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Comments" && feature.Count == 1);
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Revisions" && feature.Count >= 2);
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Content controls" && feature.Count == 1);
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "External hyperlinks"
                    && feature.Details.Any(detail => detail.Contains("https://example.com", StringComparison.OrdinalIgnoreCase)));
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Alternative format imports"
                    && feature.Details.Any(detail => detail.Contains("afchunk", StringComparison.OrdinalIgnoreCase)));
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "VBA macros"
                    && feature.Details.Any(detail => detail.Contains("vbaProject", StringComparison.OrdinalIgnoreCase)));
                Assert.Contains(report.UnsupportedFeatures, feature => feature.Name == "Digital signatures"
                    && feature.Details.Any(detail => detail.Contains("digital signature metadata", StringComparison.OrdinalIgnoreCase)));

                Assert.True(report.HasAdvancedFeatures);
                Assert.Throws<InvalidOperationException>(() => report.EnsureNoUnsupportedFeatures());
                Assert.Throws<InvalidOperationException>(() => report.EnsureNoFeatures("VBA macros"));

                string markdown = report.ToMarkdown();
                Assert.Contains("# Word Feature Report", markdown, StringComparison.Ordinal);
                Assert.Contains("| Compatibility | VBA macros |", markdown, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void Test_WordFeatureReport_AllowsSimpleEditableDocuments() {
            string filePath = Path.Combine(_directoryWithFiles, "WordFeatureReport.Simple.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Simple");
                document.AddTable(1, 1);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordFeatureReport report = document.InspectFeatures();

                Assert.Contains(report.EditableFeatures, feature => feature.Name == "Paragraphs" && feature.Count > 0);
                Assert.Contains(report.EditableFeatures, feature => feature.Name == "Tables" && feature.Count == 1);
                Assert.Empty(report.UnsupportedFeatures);
                Assert.False(report.HasAdvancedFeatures);
                Assert.Same(report, report.EnsureNoUnsupportedFeatures());
                Assert.Same(report, report.EnsureNoAdvancedFeatures());
                Assert.Empty(report.FindFeatures("Digital signatures"));
                Assert.NotEmpty(report.FindFeatures(WordFeatureSupportLevel.Editable));
            }
        }

        [Fact]
        public void Test_WordFeatureReport_SummarizesFieldRefreshReadiness() {
            string filePath = Path.Combine(_directoryWithFiles, "WordFeatureReport.FieldRefreshReadiness.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.BuiltinDocumentProperties.Creator = "Ada Lovelace";
                document.SetDocumentVariable("ClientName", "Evotec");
                document.AddParagraph("Target heading").AddBookmark("TargetBookmark");
                AddFeatureReportComplexField(document.AddParagraph("Reference: ")._paragraph, " REF ", " \"TargetBookmark\" \\h ", "Target heading", dirty: true, locked: true);
                document.AddParagraph("Page: ").AddField(WordFieldType.Page);
                document.AddParagraph("Formula: ")._paragraph.Append(BuildFeatureReportSimpleField(" = COUNT(1, 2, 3) ", "stale", locked: false));
                document.AddParagraph("TOC: ").AddField(WordFieldType.TOC);
                document.AddParagraph("Date: ").AddField(WordFieldType.Date);
                document.AddParagraph("Variable: ").AddField(WordFieldType.DocVariable, parameters: new List<string> { "\"ClientName\"" });
                document.AddParagraph("Unsupported: ")._paragraph.Append(BuildFeatureReportSimpleField(" SILLYFIELD value ", "Unknown", locked: false));
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordFeatureReport report = document.InspectFeatures();
                WordFeatureFinding fields = Assert.Single(report.FindFeatures("Fields"));

                Assert.Equal(WordFeatureSupportLevel.PartiallyEditable, fields.SupportLevel);
                Assert.Equal(7, fields.Count);
                Assert.Contains(fields.Details, detail => detail == "Simple fields: 6");
                Assert.Contains(fields.Details, detail => detail == "Complex fields: 1");
                Assert.Contains(fields.Details, detail => detail == "Deterministic refresh candidates: 5");
                Assert.Contains(fields.Details, detail => detail.Contains("Refreshable field types:", StringComparison.Ordinal)
                    && detail.Contains("Date: 1", StringComparison.Ordinal)
                    && detail.Contains("DocVariable: 1", StringComparison.Ordinal)
                    && detail.Contains("Formula: 1", StringComparison.Ordinal)
                    && detail.Contains("Page: 1", StringComparison.Ordinal)
                    && detail.Contains("Ref: 1", StringComparison.Ordinal));
                Assert.Contains(fields.Details, detail => detail == "Queued/manual refresh fields: TOC: 1");
                Assert.DoesNotContain(fields.Details, detail => detail.StartsWith("Known unsupported refresh fields", StringComparison.Ordinal));
                Assert.Contains(fields.Details, detail => detail == "Field parser diagnostics: 1");

                string markdown = report.ToMarkdown();
                Assert.Contains("Deterministic refresh candidates: 5", markdown, StringComparison.Ordinal);
                Assert.DoesNotContain("Known unsupported refresh fields", markdown, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void Test_WordFeatureReport_DetectsContentControlDataBindings() {
            string filePath = Path.Combine(_directoryWithFiles, "WordFeatureReport.ContentControlDataBinding.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Template");
                document.Save(false);
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, true)) {
                MainDocumentPart mainPart = wordDocument.MainDocumentPart!;
                CustomXmlPart customXmlPart = mainPart.AddCustomXmlPart(CustomXmlPartType.CustomXml);
                using (var stream = new MemoryStream(Encoding.UTF8.GetBytes("<root><client>Acme</client></root>"))) {
                    customXmlPart.FeedData(stream);
                }

                mainPart.Document.Body!.Append(
                    new SdtBlock(
                        new SdtProperties(
                            new SdtAlias { Val = "Client" },
                            new DocumentFormat.OpenXml.Wordprocessing.Tag { Val = "Client" },
                            new DataBinding {
                                StoreItemId = "{11111111-1111-1111-1111-111111111111}",
                                XPath = "/root[1]/client[1]"
                            }),
                        new SdtContentBlock(
                            new Paragraph(
                                new Run(
                                    new Text("Client placeholder"))))));
                mainPart.Document.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordFeatureReport report = document.InspectFeatures();
                WordFeatureFinding dataBinding = Assert.Single(report.FindFeatures("Content-control data bindings"));

                Assert.Equal(WordFeatureSupportLevel.PartiallyEditable, dataBinding.SupportLevel);
                Assert.Equal(1, dataBinding.Count);
                Assert.Contains(dataBinding.Details, detail => detail.Contains("/root[1]/client[1]", StringComparison.OrdinalIgnoreCase));

                InvalidOperationException advancedException = Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                Assert.Contains("Custom XML parts", advancedException.Message);
            }
        }

        [Fact]
        public void Test_WordFeatureReport_DetectsVisualizationAndEquationPackageDetails() {
            string filePath = Path.Combine(_directoryWithFiles, "WordFeatureReport.VisualizationAndEquationDetails.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                var chart = document.AddChart("Revenue");
                chart.AddChartAxisX(new List<string> { "Q1", "Q2" });
                chart.AddLine("Sales", new List<int> { 10, 20 }, OfficeIMO.Drawing.OfficeColor.Blue);
                document.AddSmartArt(SmartArtType.BasicProcess);
                const string omml = "<m:oMathPara xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:oMath><m:r><m:t>x=1</m:t></m:r></m:oMath></m:oMathPara>";
                document.AddEquation(omml);
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordFeatureReport report = document.InspectFeatures();

                WordFeatureFinding charts = Assert.Single(report.FindFeatures("Charts"));
                WordFeatureFinding smartArt = Assert.Single(report.FindFeatures("SmartArt"));
                WordFeatureFinding equations = Assert.Single(report.FindFeatures("Equations"));

                Assert.Equal(WordFeatureSupportLevel.PartiallyEditable, charts.SupportLevel);
                Assert.Equal(1, charts.Count);
                Assert.Contains(charts.Details, detail => detail.Contains("chart", StringComparison.OrdinalIgnoreCase));
                Assert.Equal(WordFeatureSupportLevel.Preserved, smartArt.SupportLevel);
                Assert.Equal(1, smartArt.Count);
                Assert.Contains(smartArt.Details, detail => detail.Contains("diagrams", StringComparison.OrdinalIgnoreCase));
                Assert.Equal(WordFeatureSupportLevel.PartiallyEditable, equations.SupportLevel);
                Assert.Equal(1, equations.Count);
                Assert.Contains(equations.Details, detail => detail.Contains("document.xml", StringComparison.OrdinalIgnoreCase));

                InvalidOperationException advancedException = Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                Assert.Contains("SmartArt", advancedException.Message);
            }
        }

        [Fact]
        public void Test_WordFeatureReport_DetectsDocumentAutomationMetadata() {
            string filePath = Path.Combine(_directoryWithFiles, "WordFeatureReport.DocumentAutomationMetadata.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.SetDocumentVariable("ClientName", "Acme");
                var source = new WordBibliographySource("S1", DataSourceValues.Book) {
                    Title = "Automation Notes",
                    Author = "OfficeIMO",
                    Year = "2026"
                };
                document.BibliographySources[source.Tag!] = source;
                document.AddCitation(source.Tag!);
                document.Save(false);
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, true)) {
                MainDocumentPart mainPart = wordDocument.MainDocumentPart!;
                CustomXmlPart bibliographyPart = mainPart.AddCustomXmlPart(CustomXmlPartType.Bibliography);
                new Sources(
                    new Source(
                        new SourceType { Text = DataSourceValues.Book.ToString() },
                        new DocumentFormat.OpenXml.Bibliography.Tag { Text = "S1" },
                        new Title { Text = "Automation Notes" }))
                    .Save(bibliographyPart);

                DocumentSettingsPart settingsPart = mainPart.DocumentSettingsPart ?? mainPart.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings ??= new Settings();
                settingsPart.AddExternalRelationship(
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/attachedTemplate",
                    new Uri("https://example.com/templates/report.dotx"));
                settingsPart.Settings.Save();
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordFeatureReport report = document.InspectFeatures();

                WordFeatureFinding variables = Assert.Single(report.FindFeatures("Document variables"));
                WordFeatureFinding bibliography = Assert.Single(report.FindFeatures("Bibliography sources"));
                WordFeatureFinding attachedTemplate = Assert.Single(report.FindFeatures("Attached templates"));

                Assert.Equal(WordFeatureSupportLevel.Editable, variables.SupportLevel);
                Assert.Equal(1, variables.Count);
                Assert.Equal(WordFeatureSupportLevel.Editable, bibliography.SupportLevel);
                Assert.True(bibliography.Count >= 1);
                Assert.Equal(WordFeatureSupportLevel.Preserved, attachedTemplate.SupportLevel);
                Assert.Equal(1, attachedTemplate.Count);
                Assert.Contains(attachedTemplate.Details, detail => detail.Contains("attachedTemplate", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(attachedTemplate.Details, detail => detail.Contains("report.dotx", StringComparison.OrdinalIgnoreCase));
                Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
            }
        }

        [Fact]
        public void Test_WordFeatureReport_DetectsAdvancedPackageSignals() {
            string filePath = Path.Combine(_directoryWithFiles, "WordFeatureReport.AdvancedPackageSignals.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Advanced package signals");
                document.Save(false);
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, true)) {
                MainDocumentPart mainPart = wordDocument.MainDocumentPart!;
                AddExtendedPart(mainPart,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml",
                    "<w:glossaryDocument xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:docParts /></w:glossaryDocument>");
                AddExtendedPart(mainPart,
                    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
                    "<w15:commentsEx xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" />");
                AddExtendedPart(mainPart,
                    "http://schemas.microsoft.com/office/2011/relationships/commentsIds",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsIds+xml",
                    "<w16cid:commentsIds xmlns:w16cid=\"http://schemas.microsoft.com/office/word/2016/wordml/cid\" />");
                AddExtendedPart(mainPart,
                    "http://schemas.microsoft.com/office/2011/relationships/webextension",
                    "application/vnd.ms-office.webextension+xml",
                    "<we:webextension xmlns:we=\"http://schemas.microsoft.com/office/webextensions/webextension/2010/11\" />");
                AddExtendedPart(mainPart,
                    "http://schemas.microsoft.com/office/2011/relationships/taskpanes",
                    "application/vnd.ms-office.webextensiontaskpanes+xml",
                    "<wetp:taskpanes xmlns:wetp=\"http://schemas.microsoft.com/office/webextensions/taskpanes/2010/11\" />");
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordFeatureReport report = document.InspectFeatures();

                WordFeatureFinding glossary = Assert.Single(report.FindFeatures("Building blocks and glossary"));
                WordFeatureFinding modernComments = Assert.Single(report.FindFeatures("Modern comment metadata"));
                WordFeatureFinding webExtensions = Assert.Single(report.FindFeatures("Web extensions and task panes"));

                Assert.Equal(WordFeatureSupportLevel.Preserved, glossary.SupportLevel);
                Assert.Equal(1, glossary.Count);
                Assert.Contains(glossary.Details, detail => detail.Contains("glossary", StringComparison.OrdinalIgnoreCase));
                Assert.Equal(WordFeatureSupportLevel.Preserved, modernComments.SupportLevel);
                Assert.Equal(2, modernComments.Count);
                Assert.Contains(modernComments.Details, detail => detail.Contains("commentsExtended", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(modernComments.Details, detail => detail.Contains("commentsIds", StringComparison.OrdinalIgnoreCase));
                Assert.Equal(WordFeatureSupportLevel.Preserved, webExtensions.SupportLevel);
                Assert.Equal(2, webExtensions.Count);
                Assert.Contains(webExtensions.Details, detail => detail.Contains("webextension", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(webExtensions.Details, detail => detail.Contains("taskpane", StringComparison.OrdinalIgnoreCase));

                InvalidOperationException advancedException = Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                Assert.Contains("Building blocks and glossary", advancedException.Message);
                Assert.Contains("Modern comment metadata", advancedException.Message);
                Assert.Contains("Web extensions and task panes", advancedException.Message);
            }
        }

        [Fact]
        public void Test_WordFeatureReport_DetectsExternalLinkedImages() {
            string filePath = Path.Combine(_directoryWithFiles, "WordFeatureReport.ExternalLinkedImages.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Linked logo")
                    .AddImage(new Uri("https://example.com/assets/logo.png"), 64, 32, description: "Linked logo");
                document.Save(false);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordFeatureReport report = document.InspectFeatures();

                WordFeatureFinding images = Assert.Single(report.FindFeatures("External linked images"));

                Assert.Equal(WordFeatureSupportLevel.PartiallyEditable, images.SupportLevel);
                Assert.Equal(1, images.Count);
                Assert.Contains(images.Details, detail => detail.Contains("relationships/image", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(images.Details, detail => detail.Contains("https://example.com/assets/logo.png", StringComparison.OrdinalIgnoreCase));
                Assert.Contains(report.PartiallyEditableFeatures, feature => feature.Name == "Images" && feature.Count == 1);
                Assert.Empty(report.UnsupportedFeatures);
                Assert.Same(report, report.EnsureNoUnsupportedFeatures());
            }
        }

        [Fact]
        public void Test_WordFeatureReport_DetectsActiveXControlPackageSignals() {
            string filePath = Path.Combine(_directoryWithFiles, "WordFeatureReport.ActiveXControls.docx");
            File.Delete(filePath);

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("ActiveX package signals");
                document.Save(false);
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, true)) {
                MainDocumentPart mainPart = wordDocument.MainDocumentPart!;
                AddExtendedPart(mainPart,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control",
                    "application/vnd.ms-office.activeX+xml",
                    "<ax:ocx xmlns:ax=\"http://schemas.microsoft.com/office/2006/activeX\" />");
                AddExtendedPart(mainPart,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/activeXControlBinary",
                    "application/vnd.ms-office.activeX.bin",
                    new byte[] { 1, 2, 3, 4 });
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordFeatureReport report = document.InspectFeatures();
                WordFeatureFinding activeX = Assert.Single(report.FindFeatures("ActiveX controls"));

                Assert.Equal(WordFeatureSupportLevel.Preserved, activeX.SupportLevel);
                Assert.Equal(2, activeX.Count);
                Assert.Contains(activeX.Details, detail => detail.Contains("activeX", StringComparison.OrdinalIgnoreCase));

                InvalidOperationException advancedException = Assert.Throws<InvalidOperationException>(() => report.EnsureNoAdvancedFeatures());
                Assert.Contains("ActiveX controls", advancedException.Message);
            }
        }

        [Fact]
        public void Test_WordFeatureReport_RoundTripPreservesAdvancedPackageSignals() {
            string filePath = Path.Combine(_directoryWithFiles, "WordFeatureReport.PreserveAdvancedPackageSignals.docx");
            File.Delete(filePath);

            byte[] glossaryBytes = Encoding.UTF8.GetBytes("<w:glossaryDocument xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\"><w:docParts /></w:glossaryDocument>");
            byte[] commentsExtendedBytes = Encoding.UTF8.GetBytes("<w15:commentsEx xmlns:w15=\"http://schemas.microsoft.com/office/word/2012/wordml\" />");
            byte[] webExtensionBytes = Encoding.UTF8.GetBytes("<we:webextension xmlns:we=\"http://schemas.microsoft.com/office/webextensions/webextension/2010/11\" />");
            byte[] activeXControlBytes = Encoding.UTF8.GetBytes("<ax:ocx xmlns:ax=\"http://schemas.microsoft.com/office/2006/activeX\" />");
            byte[] activeXBinaryBytes = new byte[] { 4, 3, 2, 1 };

            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Preserve package signals");
                document.Save(false);
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, true)) {
                MainDocumentPart mainPart = wordDocument.MainDocumentPart!;
                AddExtendedPart(mainPart,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document.glossary+xml",
                    glossaryBytes);
                AddExtendedPart(mainPart,
                    "http://schemas.microsoft.com/office/2011/relationships/commentsExtended",
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml",
                    commentsExtendedBytes);
                AddExtendedPart(mainPart,
                    "http://schemas.microsoft.com/office/2011/relationships/webextension",
                    "application/vnd.ms-office.webextension+xml",
                    webExtensionBytes);
                AddExtendedPart(mainPart,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/control",
                    "application/vnd.ms-office.activeX+xml",
                    activeXControlBytes);
                AddExtendedPart(mainPart,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/activeXControlBinary",
                    "application/vnd.ms-office.activeX.bin",
                    activeXBinaryBytes);
            }

            using (WordDocument document = WordDocument.Load(filePath)) {
                document.AddParagraph("OfficeIMO edit after package metadata.");
                document.Save(false);
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false)) {
                MainDocumentPart mainPart = wordDocument.MainDocumentPart!;
                AssertPartBytes(mainPart, "glossary", glossaryBytes);
                AssertPartBytes(mainPart, "commentsExtended", commentsExtendedBytes);
                AssertPartBytes(mainPart, "webextension", webExtensionBytes);
                AssertPartBytes(mainPart, "activeX+xml", activeXControlBytes);
                AssertPartBytes(mainPart, "activeX.bin", activeXBinaryBytes);
            }

            using (WordDocument document = WordDocument.Load(filePath, readOnly: true)) {
                WordFeatureReport report = document.InspectFeatures();

                Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Building blocks and glossary"
                    && feature.Details.Any(detail => detail.Contains("glossary", StringComparison.OrdinalIgnoreCase)));
                Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Modern comment metadata"
                    && feature.Details.Any(detail => detail.Contains("commentsExtended", StringComparison.OrdinalIgnoreCase)));
                Assert.Contains(report.PreservedFeatures, feature => feature.Name == "Web extensions and task panes"
                    && feature.Details.Any(detail => detail.Contains("webextension", StringComparison.OrdinalIgnoreCase)));
                Assert.Contains(report.PreservedFeatures, feature => feature.Name == "ActiveX controls"
                    && feature.Details.Any(detail => detail.Contains("activeX", StringComparison.OrdinalIgnoreCase)));
            }
        }

        private static void AddExtendedPart(OpenXmlPartContainer container, string relationshipType, string contentType, string xml) {
            AddExtendedPart(container, relationshipType, contentType, Encoding.UTF8.GetBytes(xml));
        }

        private static void AddExtendedPart(OpenXmlPartContainer container, string relationshipType, string contentType, byte[] bytes) {
            ExtendedPart part = container.AddExtendedPart(relationshipType, contentType, "xml");
            using var stream = new MemoryStream(bytes);
            part.FeedData(stream);
        }

        private static void AddFeatureReportComplexField(Paragraph paragraph, string instructionPart1, string instructionPart2, string resultText, bool dirty, bool locked) {
            paragraph.Append(
                new Run(new FieldChar { FieldCharType = FieldCharValues.Begin, Dirty = dirty, FieldLock = locked }),
                new Run(new FieldCode { Text = instructionPart1, Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldCode { Text = instructionPart2, Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
                new Run(new Text(resultText) { Space = SpaceProcessingModeValues.Preserve }),
                new Run(new FieldChar { FieldCharType = FieldCharValues.End }));
        }

        private static SimpleField BuildFeatureReportSimpleField(string instruction, string resultText, bool locked) {
            return new SimpleField(
                new Run(
                    new Text(resultText) { Space = SpaceProcessingModeValues.Preserve })) {
                Instruction = instruction,
                FieldLock = locked
            };
        }

        private static void AssertPartBytes(OpenXmlPartContainer container, string contentTypeFragment, byte[] expectedBytes) {
            OpenXmlPart part = Assert.Single(container.Parts.Select(part => part.OpenXmlPart),
                part => part.ContentType.IndexOf(contentTypeFragment, StringComparison.OrdinalIgnoreCase) >= 0);
            using var stream = part.GetStream(FileMode.Open, FileAccess.Read);
            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            Assert.Equal(expectedBytes, buffer.ToArray());
        }
    }
}
