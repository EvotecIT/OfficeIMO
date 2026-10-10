using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    /// <summary>Resolves section owners of an authored or inherited header/footer relationship.</summary>
    internal static class WordStoryLayoutResolver {
        internal static List<WordSection> GetOwners(WordDocument document, string relationshipId, bool header) {
            var owners = new List<WordSection>();
            var inheritedTypes = new HashSet<HeaderFooterValues>();
            foreach (WordSection section in document.Sections) {
                IEnumerable<HeaderFooterReferenceType> references = header
                    ? section._sectionProperties.Elements<HeaderReference>()
                    : section._sectionProperties.Elements<FooterReference>();
                foreach (HeaderFooterReferenceType reference in references) {
                    HeaderFooterValues type = reference.Type?.Value ?? HeaderFooterValues.Default;
                    if (reference.Id?.Value == relationshipId) inheritedTypes.Add(type);
                    else inheritedTypes.Remove(type);
                }
                // An omitted reference inherits the preceding section's story of that type.
                if (inheritedTypes.Count > 0) owners.Add(section);
            }
            return owners;
        }
    }
}
