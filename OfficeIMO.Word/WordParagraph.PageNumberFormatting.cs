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

    private void UpdatePageNumberFormatting<T>(T? value) where T : OpenXmlElement {
        PageNumberFormattingGroup? group = _run?.Annotation<PageNumberFormattingGroup>();
        if (group == null) return;
        foreach (Run run in group.Runs) {
            if (!run.Ancestors<Paragraph>().Any(paragraph => ReferenceEquals(paragraph, _paragraph))) continue;
            RunProperties properties = run.RunProperties ??= new RunProperties();
            properties.RemoveAllChildren<T>();
            if (value != null) properties.AddChild(value.CloneNode(true), true);
        }
    }
}
