using DocumentFormat.OpenXml.Wordprocessing;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Word.Rtf;

/// <content>Maps explicit section headers and footers while retaining inheritance through absent references.</content>
public static partial class WordRtfConverterExtensions {
    private static void ApplyFirstPageHeaderFooterSelection(WordSection section, bool enabled) {
        // Selection and declaration are separate: absent references inherit earlier stories.
        section._sectionProperties.RemoveAllChildren<TitlePage>();
        if (enabled) section._sectionProperties.Append(new TitlePage());
    }

    private static void CopySectionHeaderFooters(WordSection source, RtfSection destination, RtfDocument document,
        Dictionary<string, int> revisionAuthors) {
        CopyHeaderFooter(source.Header.Default, document, destination.AddHeader, RtfHeaderFooterKind.Header, revisionAuthors);
        CopyHeaderFooter(source.Header.First, document, destination.AddHeader, RtfHeaderFooterKind.FirstHeader, revisionAuthors);
        CopyHeaderFooter(source.Header.Even, document, destination.AddHeader, RtfHeaderFooterKind.LeftHeader, revisionAuthors);
        CopyHeaderFooter(source.Footer.Default, document, destination.AddFooter, RtfHeaderFooterKind.Footer, revisionAuthors);
        CopyHeaderFooter(source.Footer.First, document, destination.AddFooter, RtfHeaderFooterKind.FirstFooter, revisionAuthors);
        CopyHeaderFooter(source.Footer.Even, document, destination.AddFooter, RtfHeaderFooterKind.LeftFooter, revisionAuthors);
    }

    private static void ApplySectionHeaderFooters(RtfDocument source, WordDocument destination, WordSection[] sections) {
        var owned = new HashSet<RtfHeaderFooter>(source.Sections.SelectMany(section => section.HeaderFooters));
        for (int index = 0; index < sections.Length; index++) {
            IEnumerable<RtfHeaderFooter> declarations = source.Sections[index].HeaderFooters;
            if (index == 0) declarations = source.HeaderFooters.Where(header => !owned.Contains(header)).Concat(declarations);
            foreach (RtfHeaderFooter declaration in declarations) {
                WordHeaderFooter headerFooter = GetSectionHeaderFooter(destination, sections[index], declaration.Kind);
                foreach (WordParagraph existing in headerFooter.Paragraphs.ToArray()) existing.Remove();
                foreach (RtfParagraph paragraph in declaration.Paragraphs) AppendParagraph(headerFooter, paragraph, source);
            }
        }
    }

    private static WordHeaderFooter GetSectionHeaderFooter(WordDocument document, WordSection section, RtfHeaderFooterKind kind) {
        HeaderFooterValues type = kind switch {
            RtfHeaderFooterKind.FirstHeader or RtfHeaderFooterKind.FirstFooter => HeaderFooterValues.First,
            RtfHeaderFooterKind.LeftHeader or RtfHeaderFooterKind.LeftFooter => HeaderFooterValues.Even,
            _ => HeaderFooterValues.Default
        };
        bool isFooter = kind is RtfHeaderFooterKind.Footer or RtfHeaderFooterKind.FirstFooter or
            RtfHeaderFooterKind.LeftFooter or RtfHeaderFooterKind.RightFooter;
        if (isFooter) {
            WordHeadersAndFooters.AddFooterReference(document, section, type);
            return type == HeaderFooterValues.First ? section.Footer.First! :
                type == HeaderFooterValues.Even ? section.Footer.Even! : section.Footer.Default!;
        }
        WordHeadersAndFooters.AddHeaderReference(document, section, type);
        return type == HeaderFooterValues.First ? section.Header.First! :
            type == HeaderFooterValues.Even ? section.Header.Even! : section.Header.Default!;
    }
}
