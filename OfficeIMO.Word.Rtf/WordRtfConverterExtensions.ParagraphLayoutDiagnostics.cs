using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
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
        var elements = new HashSet<OpenXmlElement>();
        foreach (WordStoryRootCandidate story in storyRoots) {
            foreach (Paragraph paragraph in story.Root.Descendants<Paragraph>()) {
                if (paragraph.ParagraphProperties is { } properties) elements.Add(properties);
            }
        }
        Styles? styles = document._wordprocessingDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        if (styles != null) {
            if (styles.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle is { } defaults) {
                elements.Add(defaults);
            }
            foreach (Style style in styles.Elements<Style>()) {
                if (style.StyleParagraphProperties != null) elements.Add(style.StyleParagraphProperties);
            }
        }
        int frames = elements.Count(properties => properties.GetFirstChild<FrameProperties>() != null);
        if (frames > 0) report.Add(RtfConversionSeverity.Warning, "WordRtfParagraphFramesFlattened",
            "Word paragraph frame formatting is not represented by the RTF bridge; paragraph text remains in ordinary flow.",
            RtfConversionAction.Flattened, feature: "paragraph-frame", count: frames);
        int clearTabs = elements.Sum(properties => properties.GetFirstChild<Tabs>()?
            .Elements<TabStop>().Count(tab => tab.Val?.Value == TabStopValues.Clear) ?? 0);
        if (clearTabs > 0) report.Add(RtfConversionSeverity.Warning, "WordRtfTabClearsOmitted",
            "Word tab-stop clearing instructions have no representation in the RTF bridge and are omitted; active tab stops are converted separately.",
            RtfConversionAction.Omitted, feature: "tab-clear", count: clearTabs);
    }
}
