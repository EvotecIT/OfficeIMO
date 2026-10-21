using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordSection {
        /// <summary>
        /// Finds individual note references throughout this section's body content without changing
        /// the top-level paragraph collection. Each reference retains its source paragraph and run
        /// so that editing or removing a collected note changes only its original reference.
        /// </summary>
        private IEnumerable<(Paragraph Paragraph, Run Run, TReference Reference)> EnumerateNoteReferences<TReference>()
            where TReference : OpenXmlElement {
            int targetSection = GetSectionOrdinal();
            int finalSection = GetSectionCount() - 1;
            var body = _wordprocessingDocument?.MainDocumentPart?.Document?.Body;
            if (body == null) {
                yield break;
            }

            int currentSection = 0;
            foreach (OpenXmlElement element in body.ChildElements) {
                if (currentSection == targetSection) {
                    foreach (TReference reference in element.Descendants<TReference>()) {
                        if (reference.Parent is not Run run ||
                            run.Ancestors().Any(ancestor => ancestor is DeletedRun || ancestor is MoveFromRun)) {
                            continue;
                        }

                        Paragraph? paragraph = run.Ancestors<Paragraph>().FirstOrDefault();
                        if (paragraph != null) {
                            yield return (paragraph, run, reference);
                        }
                    }
                }

                if (element is Paragraph boundary && IsSectionBoundaryParagraph(boundary) && currentSection < finalSection) {
                    currentSection++;
                    if (currentSection > targetSection) {
                        yield break;
                    }
                }
            }
        }
    }
}
