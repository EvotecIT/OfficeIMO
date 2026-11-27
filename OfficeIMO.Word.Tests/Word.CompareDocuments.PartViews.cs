using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Office2010.Word;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.OpenXml;
using Xunit;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using WDrawing = DocumentFormat.OpenXml.Wordprocessing.Drawing;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void PartViewSharesOneViewForTwoSpellingsOfAPartName() {
            string path = Path.Combine(_directoryWithFiles, "part_view_part_name_spellings.docx");
            using (WordDocument document = WordDocument.Create(path)) {
                document.AddParagraph("Body");
                document.Save();
            }

            using WordDocument loaded = WordDocument.Load(path);
            WordPackageView package = WordPackageView.For(loaded);
            // A load that goes straight to the SDK package, as on .NET Framework, has no owned part names to compare.
            if (package.IsSdk) return;

            WordPartView plain = package.OwnedPart("/word/header1.xml");
            WordPartView escaped = package.OwnedPart("/word/%68eader1.xml");

            Assert.Same(plain, escaped);
            Assert.Equal("/word/header1.xml", escaped.PartName);
        }

        [Fact]
        public void PartViewAttributesIncludeMarkupCompatibilityAttributesOnBothBackings() {
            string path = Path.Combine(_directoryWithFiles, "part_view_markup_compatibility_attributes.docx");
            using (WordDocument document = WordDocument.Create(path)) {
                document.AddParagraph("Body");
                document.Save();
            }

            using (WordprocessingDocument package = WordprocessingDocument.Open(path, true)) {
                Paragraph paragraph = package.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First();
                paragraph.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
                paragraph.MCAttributes = new MarkupCompatibilityAttributes { Ignorable = "w14" };
                paragraph.ParagraphId = "1A2B3C4D";
                package.Save();
            }

            using WordDocument owned = WordDocument.Load(path);
            using WordDocument sdk = WordDocument.Load(path);
            _ = sdk._wordprocessingDocument;

            string[] ownedAttributes = FirstBodyParagraphAttributes(WordPackageView.For(owned));
            string[] sdkAttributes = FirstBodyParagraphAttributes(WordPackageView.For(sdk));

            Assert.Contains("http://schemas.openxmlformats.org/markup-compatibility/2006 Ignorable=w14", ownedAttributes);
            Assert.Equal(sdkAttributes, ownedAttributes);
        }

        private static string[] FirstBodyParagraphAttributes(WordPackageView package) {
            OfficeIMO.OpenXml.Internal.OfficeXmlNode paragraph = package.Main!.Root
                .Child(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.WordprocessingML, "body")
                .Child(OfficeIMO.Core.Internal.OpenXml.OfficeXmlNamespaces.WordprocessingML, "p");
            return WordXmlNodes.Attributes(paragraph)
                .Select(attribute => attribute.NamespaceUri + " " + attribute.LocalName + "=" + attribute.Value)
                .OrderBy(text => text, StringComparer.Ordinal)
                .ToArray();
        }

        [Fact]
        public void DrawingLayoutIgnoresGroupGraphicDataInsideExtensionLists() {
            var drawing = new WDrawing(
                "<w:drawing xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" " +
                "xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" " +
                "xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                "<wp:inline><wp:extent cx=\"914400\" cy=\"457200\"/><wp:docPr id=\"1\" name=\"Picture 1\">" +
                "<a:extLst><a:ext uri=\"{00000000-0000-0000-0000-000000000001}\">" +
                "<a:graphicData uri=\"http://schemas.microsoft.com/office/word/2010/wordprocessingGroup\"/>" +
                "</a:ext></a:extLst></wp:docPr>" +
                "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\"/></a:graphic>" +
                "</wp:inline></w:drawing>");

            Assert.True(WordDrawingLayoutReader.TryRead(drawing, out WordDrawingLayoutInfo layout));
            Assert.False(layout.IsGroup);
        }

        [Fact]
        public void CompareStructureInPlaceRedlineKeepsRunLevelControlsInConflictRevisions() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_conflict_revision_control_source.docx");
            CreateControlInConflictRevision(sourcePath, "Contoso");
            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_conflict_revision_control_target.docx");
            CreateControlInConflictRevision(targetPath, "Fabrikam");

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_conflict_revision_control_output.docx");
            WordComparisonResult result = WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 11, 27, 9, 0, 0, DateTimeKind.Utc),
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new HashSet<WordComparisonScope> {
                            WordComparisonScope.ContentControl
                        }
                    }
                });

            Assert.Single(result.Findings, finding =>
                finding.Scope == WordComparisonScope.ContentControl &&
                finding.ChangeKind == WordComparisonChangeKind.Modified);

            using WordprocessingDocument redline = WordprocessingDocument.Open(outputPath, false);
            SdtRun contentControl = Assert.Single(redline.MainDocumentPart!.Document!.Body!.Descendants<RunConflictInsertion>().Single().Elements<SdtRun>());
            Assert.DoesNotContain(contentControl.Descendants(), element => element.LocalName == "p");
            Assert.Contains(contentControl.Descendants<DeletedRun>(), run => run.InnerText == "Contoso");
            Assert.Contains(contentControl.Descendants<InsertedRun>(), run => run.InnerText == "Fabrikam");
        }

        private static void CreateControlInConflictRevision(string path, string text) {
            using (WordDocument document = WordDocument.Create(path)) {
                document.AddParagraph("Client: ").AddStructuredDocumentTag(text, "Client", "ClientName");
                document.Save();
            }

            using WordprocessingDocument package = WordprocessingDocument.Open(path, true);
            SdtRun control = package.MainDocumentPart!.Document!.Body!.Descendants<SdtRun>().Single();
            var conflict = new RunConflictInsertion { Id = "1", Author = "Reviewer" };
            control.InsertBeforeSelf(conflict);
            control.Remove();
            conflict.Append(control);
            package.Save();
        }

        [Fact]
        public void CompareStructureInPlaceRedlineCopiesImageRelationshipsInsideDrawingExtensionLists() {
            string sourcePath = Path.Combine(_directoryWithFiles, "compare_redline_image_extension_blip_source.docx");
            using (WordDocument document = WordDocument.Create(sourcePath)) {
                document.AddParagraph("Before image");
                document.AddParagraph().AddImage(Path.Combine(_directoryWithImages, "EvotecLogo.png"), OfficeIMO.OfficeLength.Pixels(80), OfficeIMO.OfficeLength.Pixels(40));
                document.Save();
            }

            string sourceRelationshipId;
            using (WordprocessingDocument package = WordprocessingDocument.Open(sourcePath, true)) {
                WDrawing drawing = package.MainDocumentPart!.Document!.Body!.Descendants<WDrawing>().Single();
                sourceRelationshipId = drawing.Descendants<A.Blip>().Single().Embed!.Value!;
                var nestedBlip = new OpenXmlUnknownElement("a", "blip", "http://schemas.openxmlformats.org/drawingml/2006/main");
                nestedBlip.SetAttribute(new OpenXmlAttribute("r", "embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", sourceRelationshipId));
                var extension = new A.NonVisualDrawingPropertiesExtension { Uri = "{00000000-0000-0000-0000-000000000002}" };
                extension.Append(nestedBlip);
                drawing.Descendants<DW.DocProperties>().Single().Append(new A.NonVisualDrawingPropertiesExtensionList(extension));
                package.Save();
            }

            string targetPath = Path.Combine(_directoryWithFiles, "compare_redline_image_extension_blip_target.docx");
            using (WordDocument document = WordDocument.Create(targetPath)) {
                document.AddParagraph("Before image");
                document.Save();
            }

            string outputPath = Path.Combine(_directoryWithFiles, "compare_redline_image_extension_blip_output.docx");
            WordDocumentComparer.CreateRedlineDocument(
                sourcePath,
                targetPath,
                outputPath,
                new WordComparisonRedlineOptions {
                    Mode = WordComparisonRedlineMode.InPlaceTarget,
                    Author = "OfficeIMO Tests",
                    DateTime = new DateTime(2026, 11, 27, 9, 30, 0, DateTimeKind.Utc),
                    ComparisonOptions = new WordComparisonOptions {
                        IncludedScopes = new HashSet<WordComparisonScope> {
                            WordComparisonScope.Image
                        }
                    }
                });

            using WordprocessingDocument redline = WordprocessingDocument.Open(outputPath, false);
            MainDocumentPart mainPart = redline.MainDocumentPart!;
            DeletedRun deletedImage = Assert.Single(mainPart.Document!.Body!.Descendants<DeletedRun>(), run => run.Descendants<WDrawing>().Any());
            OpenXmlElement copiedNestedBlip = Assert.Single(
                deletedImage.Descendants<DW.DocProperties>().Single().Descendants(),
                element => element.LocalName == "blip");
            string copiedRelationshipId = copiedNestedBlip.GetAttribute("embed", "http://schemas.openxmlformats.org/officeDocument/2006/relationships").Value!;
            Assert.NotEqual(sourceRelationshipId, copiedRelationshipId);
            Assert.IsType<ImagePart>(mainPart.GetPartById(copiedRelationshipId));
        }
    }
}
