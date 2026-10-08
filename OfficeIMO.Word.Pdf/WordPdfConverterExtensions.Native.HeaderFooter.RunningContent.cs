using DocumentFormat.OpenXml;
using System.Collections.Generic;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf;

public static partial class WordPdfConverterExtensions {
    private static bool HasNativeRunningTable(WordHeaderFooter? story) => story != null &&
        story.ChildElements.Any(element => element is W.Table || element.Descendants<W.Table>().Any());

    private static bool UsesNativeRunningHeader(WordSection section) =>
        HasNativeRunningTable(section.Header?.Default) ||
        (section.DifferentFirstPage && HasNativeRunningTable(section.Header?.First)) ||
        (section.DocumentOddEvenSettingEnabled && HasNativeRunningTable(section.Header?.Even));

    private static bool UsesNativeRunningFooter(WordSection section) =>
        HasNativeRunningTable(section.Footer?.Default) ||
        (section.DifferentFirstPage && HasNativeRunningTable(section.Footer?.First)) ||
        (section.DocumentOddEvenSettingEnabled && HasNativeRunningTable(section.Footer?.Even));

    private static void ConfigureNativeRunningHeaderFooter(PdfCore.PdfPageBuilder page, WordSection section,
        WordToPdfOptions? options, NativeFontMap fontMap,
        IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeDocumentDefaults defaults) {
        if (UsesNativeRunningHeader(section)) {
            double distance = section.Margins.HeaderDistance / 20D;
            page.Header(header => {
                header.Content(CreateNativeRunningStory(section.Header?.Default, section, options, fontMap, listMarkers, defaults, false), distance);
                if (section.DifferentFirstPage)
                    header.FirstPageContent(CreateNativeRunningStory(section.Header?.First, section, options, fontMap, listMarkers, defaults, false), distance);
                if (section.DocumentOddEvenSettingEnabled)
                    header.EvenPagesContent(CreateNativeRunningStory(section.Header?.Even, section, options, fontMap, listMarkers, defaults, false), distance);
            });
        }
        if (UsesNativeRunningFooter(section)) {
            double distance = section.Margins.FooterDistance / 20D;
            page.Footer(footer => {
                footer.Content(CreateNativeRunningStory(section.Footer?.Default, section, options, fontMap, listMarkers, defaults, true), distance);
                if (section.DifferentFirstPage)
                    footer.FirstPageContent(CreateNativeRunningStory(section.Footer?.First, section, options, fontMap, listMarkers, defaults, true), distance);
                if (section.DocumentOddEvenSettingEnabled)
                    footer.EvenPagesContent(CreateNativeRunningStory(section.Footer?.Even, section, options, fontMap, listMarkers, defaults, true), distance);
            });
        }
    }

    private static Func<PdfCore.PdfRunningContentContext, Action<PdfCore.PdfContentBuilder>> CreateNativeRunningStory(
        WordHeaderFooter? story, WordSection section, WordToPdfOptions? options, NativeFontMap fontMap,
        IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeDocumentDefaults defaults, bool isFooter) {
        IReadOnlyList<WordElement> elements = story == null ? Array.Empty<WordElement>() : CollapseNativeParagraphElements(story.Elements);
        bool addPageNumber = isFooter && options?.IncludePageNumbers == true &&
            (story == null || GetNativeHeaderFooterText(story, listMarkers, fontMap, options)?.HasPageTokens != true);
        return context => content => {
            options?.CancellationToken.ThrowIfCancellationRequested();
            var flow = new NativeSpacingCollapseFlow(new NativePdfColumnFlow(content,
                new PdfCore.PageSize(context.PageWidth, context.PageHeight)));
            NativeDocumentDefaults runningDefaults = defaults with { RunningContentContext = context };
            var notes = new NativeNoteNumbering(section._document, options);
            notes.BeginSection(section);
            var destinations = new Dictionary<W.Paragraph, string>();
            for (int index = 0; index < elements.Count; index++) {
                options?.CancellationToken.ThrowIfCancellationRequested();
                if (TryRenderNativeJoinedParagraphs(flow, elements, ref index,
                    paragraph => listMarkers.TryGetValue(paragraph, out var joinedMarker) ? joinedMarker : null,
                    notes, options, runningDefaults, fontMap)) continue;
                RenderNativeElement(flow, elements[index], section,
                    paragraph => listMarkers.TryGetValue(paragraph, out var marker) ? marker : null,
                    Array.Empty<int>(), notes, options, Array.Empty<NativeTableOfContentsEntry>(), destinations,
                    context.ContentWidth, runningDefaults, fontMap,
                    index + 1 < elements.Count ? elements[index + 1] : null);
            }
            if (addPageNumber)
                content.Paragraph(paragraph => paragraph.Text(FormatNativeRunningPageTokens(
                    GetNativePageNumberFormat(options), context, null)), PdfCore.PdfAlign.Right);
        };
    }

    private static List<WordParagraph> GetNativeRunningContentRuns(WordParagraph paragraph, PdfCore.PdfRunningContentContext? context) {
        if (context == null || paragraph._paragraph == null) return GetNativeRuns(paragraph);
        var serialized = new List<(W.Run Run, string Text, bool IsField, OpenXmlElement? SourceChild, PdfCore.PdfPageNumberStyle? FieldStyle)>();
        if (!TryBuildNativeHeaderFooterParagraphText(paragraph, out _, out _, serialized)) return GetNativeRuns(paragraph);
        var runs = new List<WordParagraph>(serialized.Count);
        foreach (var item in serialized) {
            var source = new WordParagraph(paragraph._document,
                item.Run.Ancestors<W.Paragraph>().FirstOrDefault() ?? paragraph._paragraph, item.Run) {
                _hyperlink = item.Run.Ancestors<W.Hyperlink>().FirstOrDefault()
            };
            if (item.IsField) {
                string text = FormatNativeRunningPageTokens(item.Text, context, item.FieldStyle);
                runs.Add(CreateNativeRunContentView(source, new OpenXmlElement[] {
                    new W.Text(text) { Space = SpaceProcessingModeValues.Preserve }
                }));
            } else if (item.SourceChild != null) {
                runs.Add(CreateNativeRunContentView(source, new[] { item.SourceChild }));
            }
        }
        return runs;
    }

    private static string FormatNativeRunningPageTokens(string text, PdfCore.PdfRunningContentContext context, PdfCore.PdfPageNumberStyle? style) =>
        text.Replace("{page}", context.FormatPageNumber(context.PageNumber, style))
            .Replace("{pages}", context.FormatPageNumber(context.TotalPages, style))
            .Replace("{documentpages}", context.FormatPageNumber(context.DocumentPages, style));
}
