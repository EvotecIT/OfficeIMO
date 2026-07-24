using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class Word {
        [Fact]
        public void InspectionSnapshot_BoundsRecursiveFootnoteAndEndnoteReferences() {
            string filePath = Path.Combine(_directoryWithFiles, "InspectionRecursiveNotes.docx");
            using (WordDocument document = WordDocument.Create(filePath)) {
                document.AddParagraph("Footnote anchor").AddFootNote("Footnote body");
                document.AddParagraph("Endnote anchor").AddEndNote("Endnote body");
                document.Save();
            }

            using (WordprocessingDocument package = WordprocessingDocument.Open(filePath, true)) {
                Footnote footnote = package.MainDocumentPart!.FootnotesPart!.Footnotes!
                    .Elements<Footnote>()
                    .Single(note => note.Id?.Value > 0);
                Paragraph footnoteBody = footnote.Elements<Paragraph>().Last();
                footnoteBody.Append(new Run(new Text(" recursive "), new FootnoteReference { Id = footnote.Id }));
                package.MainDocumentPart.FootnotesPart.Footnotes.Save();

                Endnote endnote = package.MainDocumentPart.EndnotesPart!.Endnotes!
                    .Elements<Endnote>()
                    .Single(note => note.Id?.Value > 0);
                Paragraph endnoteBody = endnote.Elements<Paragraph>().Last();
                endnoteBody.Append(new Run(new Text(" recursive "), new EndnoteReference { Id = endnote.Id }));
                package.MainDocumentPart.EndnotesPart.Endnotes.Save();
            }

            using WordDocument loaded = WordDocument.Load(filePath);
            WordDocumentSnapshot snapshot = loaded.CreateInspectionSnapshot();

            Assert.Single(snapshot.Sections);
            var runs = snapshot.Sections[0].Elements
                .OfType<WordParagraphSnapshot>()
                .SelectMany(paragraph => paragraph.Runs)
                .ToList();
            Assert.Contains(runs, run => run.Footnote != null);
            Assert.Contains(runs, run => run.Endnote != null);
        }
    }
}
