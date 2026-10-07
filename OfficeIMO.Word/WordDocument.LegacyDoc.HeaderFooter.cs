using OfficeIMO.Word.LegacyDoc;
using OfficeIMO.Word.LegacyDoc.Diagnostics;
using OfficeIMO.Word.LegacyDoc.Model;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.ExtendedProperties;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    public partial class WordDocument {
        private static void AddLegacyDocHeaderFooterStories(WordDocument document, IReadOnlyList<LegacyDocHeaderFooterStory> stories, LegacyDocStyleSheet styleSheet) {
            foreach (LegacyDocHeaderFooterStory story in stories) {
                if (story.SectionIndex < 0 || story.SectionIndex >= document.Sections.Count || story.Blocks.Count == 0) {
                    continue;
                }

                WordSection section = document.Sections[story.SectionIndex];
                WordHeaderFooter target = story.IsHeader
                    ? section.GetOrCreateHeader(story.Type.ToOfficeEnum())
                    : section.GetOrCreateFooter(story.Type.ToOfficeEnum());
                foreach (WordParagraph paragraph in target.Paragraphs.ToList()) {
                    paragraph.Remove();
                }

                foreach (LegacyDocBodyBlock block in story.Blocks) {
                    if (block is LegacyDocTableBlock tableBlock) {
                        AddLegacyDocTableCore((rows, columns) => target.AddTable(rows, columns, WordTableStyle.TableNormal),
                            tableBlock, styleSheet, LegacyDocNoteProjection.Empty);
                        continue;
                    }
                    if (block is not LegacyDocParagraphBlock sourceParagraph) continue;
                    WordParagraph paragraph = target.AddParagraph(sourceParagraph.Bookmarks.Count == 0 ? string.Concat(sourceParagraph.Runs.Select(run => run.Text)) : string.Empty);
                    if (sourceParagraph.Bookmarks.Count > 0) {
                        paragraph._paragraph.RemoveAllChildren<Run>();
                    }

                    ApplyLegacyDocParagraphFormatting(paragraph, sourceParagraph.Format, styleSheet);
                    ReplaceLegacyDocParagraphRuns(
                        paragraph,
                        sourceParagraph.Runs,
                        LegacyDocNoteProjection.Empty,
                        LegacyDocBookmarkProjection.Create(sourceParagraph.Bookmarks, sourceParagraph.StartCharacter, sourceParagraph.EndCharacter));
                }
            }
        }

    }
}
