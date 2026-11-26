using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    // The association belongs to the source run, so other wrappers of that run
    // retain run-scoped behavior. It is authoring state, not a DOCX extension.
    private sealed class PageNumberFormattingGroup {
        internal List<Run> Runs { get; } = new();
    }

    private void InheritPageNumberFormatting(IEnumerable<Run> runs) {
        Run source = VerifyRun();
        PageNumberFormattingGroup? group = source.Annotation<PageNumberFormattingGroup>();
        if (group == null) {
            group = new PageNumberFormattingGroup();
            source.AddAnnotation(group);
        }
        foreach (Run run in runs) {
            RunProperties properties = run.RunProperties ??= new RunProperties();
            foreach (OpenXmlElement property in source.RunProperties?.ChildElements ?? Enumerable.Empty<OpenXmlElement>()) {
                foreach (OpenXmlElement existing in properties.ChildElements.Where(child => child.GetType() == property.GetType()).ToArray()) {
                    existing.Remove();
                }
                properties.AddChild(property.CloneNode(true), true);
            }
            group.Runs.Add(run);
        }
    }

    // Copies a run property the paragraph just set, named by its element, to the runs of a page number field built from
    // this run; a property the paragraph removed is removed there too.
    private void SdkSyncPageNumberFormatting(object properties, string localName) {
        PageNumberFormattingGroup? group = _run?.Annotation<PageNumberFormattingGroup>();
        if (group == null) return;
        OpenXmlElement? value = ((OpenXmlElement)properties).ChildElements.FirstOrDefault(child => IsWordProperty(child, localName));
        foreach (Run run in group.Runs) {
            if (!run.Ancestors<Paragraph>().Any(paragraph => ReferenceEquals(paragraph, _paragraph))) continue;
            RunProperties target = run.RunProperties ??= new RunProperties();
            foreach (OpenXmlElement existing in target.ChildElements.Where(child => IsWordProperty(child, localName)).ToArray()) existing.Remove();
            if (value != null) target.AddChild(value.CloneNode(true), true);
        }
    }

    private static bool IsWordProperty(OpenXmlElement element, string localName) =>
        element.LocalName == localName && element.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
}
