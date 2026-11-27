using OfficeIMO.Rtf.Writing;

namespace OfficeIMO.Word.Rtf;

public static partial class WordRtfConverterExtensions {
    /// <summary>Reports frame formatting that the bridge replaces with ordinary paragraph flow.</summary>
    private static void AddRtfParagraphFrameDiagnostics(RtfDocument document, RtfConversionReport report) {
        int count = RtfDocumentWriter.EnumerateParagraphs(document)
            .Count(paragraph => document.GetParagraphFormatting(paragraph).Frame.HasAnyValue);
        if (count == 0) return;
        report.Add(RtfConversionSeverity.Warning, "RtfWordParagraphFramesFlattened",
            "RTF paragraph frames are converted to ordinary Word paragraph flow; frame positioning, dimensions and wrapping are not preserved.",
            RtfConversionAction.Flattened, feature: "paragraph-frame", count: count);
    }

    /// <summary>Reports unsupported Word paragraph layout metadata once per source element.</summary>
    private static void AddWordParagraphLayoutDiagnostics(WordDocument document,
        IEnumerable<WordStoryRootCandidate> storyRoots, RtfConversionReport report) {
        var elements = new HashSet<OfficeXmlNode>();
        foreach (WordStoryRootCandidate story in storyRoots) {
            foreach (OfficeXmlNode paragraph in story.Root.ModelDescendants("p")) {
                OfficeXmlNode properties = paragraph.WChild("pPr");
                if (properties.Exists) elements.Add(properties);
            }
        }
        OfficeXmlNode styles = document.StylesNode();
        if (styles.Exists) {
            OfficeXmlNode defaults = styles.WChild("docDefaults").WChild("pPrDefault").WChild("pPr");
            if (defaults.Exists) {
                elements.Add(defaults);
            }
            foreach (OfficeXmlNode style in styles.WElements("style")) {
                OfficeXmlNode properties = style.WChild("pPr");
                if (properties.Exists) elements.Add(properties);
            }
        }
        int frames = elements.Count(properties => properties.WChild("framePr").Exists);
        if (frames > 0) report.Add(RtfConversionSeverity.Warning, "WordRtfParagraphFramesFlattened",
            "Word paragraph frame formatting is not represented by the RTF bridge; paragraph text remains in ordinary flow.",
            RtfConversionAction.Flattened, feature: "paragraph-frame", count: frames);
        int clearTabs = elements.Sum(properties => properties.WChild("tabs")
            .WElements("tab").Count(tab => tab.WVal() == "clear"));
        if (clearTabs > 0) report.Add(RtfConversionSeverity.Warning, "WordRtfTabClearsOmitted",
            "Word tab-stop clearing instructions have no representation in the RTF bridge and are omitted; active tab stops are converted separately.",
            RtfConversionAction.Omitted, feature: "tab-clear", count: clearTabs);
    }
}
