using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void Test_MailMerge_RepeatingBlockRegionsPreserveSectionBreakProperties() {
            string filePath = Path.Combine(_directoryWithFiles, "MailMergeSectionRepeatingBlocks.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Report");
                document.AddParagraph("{{#each Sections}}");
                document.AddParagraph("Section: ")
                    .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"SectionTitle\"" });
                AppendSectionBreakParagraph(document, PageOrientationValues.Landscape, "1440", "2880");
                document.AddParagraph("{{/each Sections}}");
                document.AddParagraph("Appendix");

                int generated = WordMailMerge.ExecuteRepeatingBlocks(
                    document,
                    new Dictionary<string, IEnumerable<IDictionary<string, string>>> {
                        ["Sections"] = new[] {
                            new Dictionary<string, string> {
                                ["SectionTitle"] = "Scope"
                            },
                            new Dictionary<string, string> {
                                ["SectionTitle"] = "Findings"
                            }
                        }
                    });

                Assert.Equal(2, generated);
                document.Save();
            }

            using WordprocessingDocument wordDocument = WordprocessingDocument.Open(filePath, false);
            Body body = wordDocument.MainDocumentPart!.Document.Body!;
            Assert.Contains("Report", body.InnerText);
            Assert.Contains("Section: Scope", body.InnerText);
            Assert.Contains("Section: Findings", body.InnerText);
            Assert.Contains("Appendix", body.InnerText);
            Assert.DoesNotContain("{{#each Sections}}", body.InnerText);
            Assert.DoesNotContain("{{/each Sections}}", body.InnerText);
            Assert.DoesNotContain("MERGEFIELD", body.InnerXml);

            List<SectionProperties> sectionProperties = body.Descendants<SectionProperties>().ToList();
            Assert.True(sectionProperties.Count >= 2);
            Assert.True(
                sectionProperties.Take(2).All(properties =>
                    properties.GetFirstChild<SectionType>()?.Val?.Value == SectionMarkValues.NextPage
                    && properties.GetFirstChild<PageSize>()?.Orient?.Value == PageOrientationValues.Landscape
                    && properties.GetFirstChild<PageMargin>()?.Left?.Value == 1440
                    && properties.GetFirstChild<PageMargin>()?.Right?.Value == 2880));
        }

        [Fact]
        public void Test_MailMerge_ConditionalBlocksCanKeepOrRemoveSectionRegions() {
            string includedPath = Path.Combine(_directoryWithFiles, "MailMergeSectionConditionalIncluded.docx");
            CreateConditionalSectionTemplate(includedPath);
            using (WordDocument document = WordDocument.Load(includedPath)) {
                int processed = WordMailMerge.ExecuteConditionalBlocks(
                    document,
                    new Dictionary<string, bool> {
                        ["IncludeLandscapeAppendix"] = true
                    });
                WordMailMerge.Execute(
                    document,
                    new Dictionary<string, string> {
                        ["AppendixTitle"] = "Risk appendix"
                    });

                Assert.Equal(1, processed);
                document.Save();
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(includedPath, false)) {
                Body body = wordDocument.MainDocumentPart!.Document.Body!;
                Assert.Contains("Appendix: Risk appendix", body.InnerText);
                Assert.DoesNotContain("{{#IncludeLandscapeAppendix}}", body.InnerText);
                Assert.DoesNotContain("{{/IncludeLandscapeAppendix}}", body.InnerText);
                Assert.Contains(body.Descendants<SectionProperties>(), properties =>
                    properties.GetFirstChild<PageSize>()?.Orient?.Value == PageOrientationValues.Landscape);
            }

            string removedPath = Path.Combine(_directoryWithFiles, "MailMergeSectionConditionalRemoved.docx");
            CreateConditionalSectionTemplate(removedPath);
            using (WordDocument document = WordDocument.Load(removedPath)) {
                int processed = WordMailMerge.ExecuteConditionalBlocks(
                    document,
                    new Dictionary<string, bool> {
                        ["IncludeLandscapeAppendix"] = false
                    });

                Assert.Equal(1, processed);
                document.Save();
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(removedPath, false)) {
                Body body = wordDocument.MainDocumentPart!.Document.Body!;
                Assert.Contains("Contract", body.InnerText);
                Assert.Contains("Closing", body.InnerText);
                Assert.DoesNotContain("Appendix:", body.InnerText);
                Assert.DoesNotContain("{{#IncludeLandscapeAppendix}}", body.InnerText);
                Assert.DoesNotContain("{{/IncludeLandscapeAppendix}}", body.InnerText);
                Assert.DoesNotContain(body.Descendants<SectionProperties>(), properties =>
                    properties.GetFirstChild<PageSize>()?.Orient?.Value == PageOrientationValues.Landscape);
            }
        }

        [Fact]
        public void Test_MailMerge_WordAuthoredMultiSectionTemplateCanBePreflightedAndBound() {
            string sourcePath = GetFixtureDoc(Path.Combine("Word", "PremiumGaps", "TemplateMailMerge", "word-authored-multi-section-template.docx"));
            Assert.True(File.Exists(sourcePath), $"Missing Word-authored multi-section template fixture: {sourcePath}");

            string includedPath = Path.Combine(_directoryWithFiles, "MailMergeWordAuthoredMultiSectionIncluded.docx");
            File.Copy(sourcePath, includedPath, overwrite: true);
            using (WordDocument document = WordDocument.Load(includedPath)) {
                WordTemplatePreflightReport report = WordMailMerge.PreflightTemplate(
                    document,
                    mergeFieldNames: new[] { "ClientName", "AppendixTitle" },
                    conditionNames: new[] { "IncludeLandscapeAppendix" });

                Assert.True(report.CanBindTemplate);
                Assert.Equal(2, report.MergeFieldCount);
                Assert.Equal(1, report.ConditionalBlockCount);
                Assert.Equal(0, report.IssueCount);

                int processed = WordMailMerge.ExecuteConditionalBlocks(
                    document,
                    new Dictionary<string, bool> {
                        ["IncludeLandscapeAppendix"] = true
                    });
                WordMailMerge.Execute(
                    document,
                    new Dictionary<string, string> {
                        ["ClientName"] = "Northwind Traders",
                        ["AppendixTitle"] = "Risk appendix"
                    });

                Assert.Equal(1, processed);
                document.Save();
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(includedPath, false)) {
                Body body = wordDocument.MainDocumentPart!.Document.Body!;
                Assert.Contains("Contract", body.InnerText);
                Assert.Contains("Client: Northwind Traders", body.InnerText);
                Assert.Contains("Appendix: Risk appendix", body.InnerText);
                Assert.Contains("Closing", body.InnerText);
                Assert.DoesNotContain("{{#IncludeLandscapeAppendix}}", body.InnerText);
                Assert.DoesNotContain("{{/IncludeLandscapeAppendix}}", body.InnerText);
                Assert.DoesNotContain("MERGEFIELD", body.InnerXml);
                Assert.Contains(body.Descendants<SectionProperties>(), properties =>
                    properties.GetFirstChild<PageSize>()?.Orient?.Value == PageOrientationValues.Landscape &&
                    properties.GetFirstChild<PageMargin>()?.Right?.Value == 2880);
            }

            string removedPath = Path.Combine(_directoryWithFiles, "MailMergeWordAuthoredMultiSectionRemoved.docx");
            File.Copy(sourcePath, removedPath, overwrite: true);
            using (WordDocument document = WordDocument.Load(removedPath)) {
                int processed = WordMailMerge.ExecuteConditionalBlocks(
                    document,
                    new Dictionary<string, bool> {
                        ["IncludeLandscapeAppendix"] = false
                    });
                WordMailMerge.Execute(
                    document,
                    new Dictionary<string, string> {
                        ["ClientName"] = "Northwind Traders"
                    });

                Assert.Equal(1, processed);
                document.Save();
            }

            using (WordprocessingDocument wordDocument = WordprocessingDocument.Open(removedPath, false)) {
                Body body = wordDocument.MainDocumentPart!.Document.Body!;
                Assert.Contains("Contract", body.InnerText);
                Assert.Contains("Client: Northwind Traders", body.InnerText);
                Assert.Contains("Closing", body.InnerText);
                Assert.DoesNotContain("Appendix:", body.InnerText);
                Assert.DoesNotContain("{{#IncludeLandscapeAppendix}}", body.InnerText);
                Assert.DoesNotContain("{{/IncludeLandscapeAppendix}}", body.InnerText);
                Assert.DoesNotContain("MERGEFIELD", body.InnerXml);
                Assert.DoesNotContain(body.Descendants<SectionProperties>(), properties =>
                    properties.GetFirstChild<PageSize>()?.Orient?.Value == PageOrientationValues.Landscape);
            }
        }

        private static void CreateConditionalSectionTemplate(string filePath) {
            using WordDocument document = WordDocument.Create(filePath);
            document.AddParagraph("Contract");
            document.AddParagraph("{{#IncludeLandscapeAppendix}}");
            document.AddParagraph("Appendix: ")
                .AddField(WordFieldType.MergeField, parameters: new List<string> { "\"AppendixTitle\"" });
            AppendSectionBreakParagraph(document, PageOrientationValues.Landscape, "1440", "1440");
            document.AddParagraph("{{/IncludeLandscapeAppendix}}");
            document.AddParagraph("Closing");
            document.Save();
        }

        private static void AppendSectionBreakParagraph(WordDocument document, PageOrientationValues orientation, string leftMargin, string rightMargin) {
            var paragraph = new Paragraph(
                new ParagraphProperties(
                    new SectionProperties(
                        new SectionType { Val = SectionMarkValues.NextPage },
                        new PageSize {
                            Width = 15840U,
                            Height = 12240U,
                            Orient = orientation
                        },
                        new PageMargin {
                            Top = 1440,
                            Bottom = 1440,
                            Left = UInt32Value.FromUInt32(uint.Parse(leftMargin, System.Globalization.CultureInfo.InvariantCulture)),
                            Right = UInt32Value.FromUInt32(uint.Parse(rightMargin, System.Globalization.CultureInfo.InvariantCulture))
                        })));

            document._document.MainDocumentPart!.Document.Body!.Append(paragraph);
        }
    }
}
