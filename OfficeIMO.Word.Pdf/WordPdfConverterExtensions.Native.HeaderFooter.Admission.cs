using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {
    public static partial class WordPdfConverterExtensions {
        private static void PrepareNativeRunningStoryAdmission(WordDocument document, WordToPdfOptions? options,
            PdfCore.PdfOptions pdfOptions, NativeFontMap fontMap,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers, NativeDocumentDefaults defaults) {
            if (options == null) return;
            var fits = new Dictionary<(WordSection Section, bool Footer), bool>();
            WordToPdfOptions probeOptions = options.CloneForConversion();
            foreach (WordSection section in document.Sections) {
                options.CancellationToken.ThrowIfCancellationRequested();
                bool headerEligible = UsesNativeRunningHeader(section, options);
                bool footerEligible = UsesNativeRunningFooter(section, options);
                WordHeaderFooter?[] headers = { section.Header?.Default,
                    section.DifferentFirstPage ? section.Header?.First : null,
                    section.DocumentOddEvenSettingEnabled ? section.Header?.Even : null };
                WordHeaderFooter?[] footers = { section.Footer?.Default,
                    section.DifferentFirstPage ? section.Footer?.First : null,
                    section.DocumentOddEvenSettingEnabled ? section.Footer?.Even : null };
                // Table and SECTIONPAGES stories retain their required native flow.
                // This admission check preserves the simple-story fallback for ordinary content.
                bool headerOrdinary = headerEligible && !headers.Any(RequiresNativeRunningStory);
                bool footerOrdinary = footerEligible && !footers.Any(RequiresNativeRunningStory);
                double? headerHeight = headerOrdinary ? MeasureStory(headers, false) : 0D;
                double? footerHeight = footerOrdinary ? MeasureStory(footers, true) : 0D;
                PdfCore.PageSize page = GetNativePageSize(section, options);
                PdfCore.PageMargins margins = GetNativeMargins(section, options, (0D, 0D));
                double headerDistance = section.Margins.HeaderDistance / 20D;
                double footerDistance = section.Margins.FooterDistance / 20D;
                bool headerFits = headerHeight.HasValue && headerHeight.Value + headerDistance < page.Height;
                bool footerFits = footerHeight.HasValue && footerHeight.Value + footerDistance < page.Height;
                if (headerOrdinary && footerOrdinary && headerFits && footerFits &&
                    Math.Max(margins.Top, headerHeight!.Value + headerDistance) +
                    Math.Max(margins.Bottom, footerHeight!.Value + footerDistance) >= page.Height) {
                    headerFits = footerFits = false;
                }
                fits[(section, false)] = !headerOrdinary || headerFits;
                fits[(section, true)] = !footerOrdinary || footerFits;

                double? MeasureStory(IEnumerable<WordHeaderFooter?> stories, bool footer) {
                    double maximum = 0D;
                    double distance = (footer ? section.Margins.FooterDistance : section.Margins.HeaderDistance) / 20D;
                    PdfCore.PageSize size = GetNativePageSize(section, options);
                    if (distance >= size.Height) return null;
                    PdfCore.PageMargins physicalMargins = GetNativeMargins(section, options, (0D, 0D));
                    double width = size.Width - physicalMargins.Left - physicalMargins.Right;
                    PdfCore.PdfOptions frame = pdfOptions.CreateRunningContentFrame(distance, 0D);
                    frame.PageSize = size;
                    frame.MarginLeft = physicalMargins.Left;
                    frame.MarginRight = physicalMargins.Right;
                    foreach (WordHeaderFooter? story in stories.Where(story => story != null).Distinct()) {
                        options.CancellationToken.ThrowIfCancellationRequested();
                        var context = new PdfCore.PdfRunningContentContext(1, 1, 1, 1, 1, 1,
                            width, size.Width, size.Height, frame.PageNumberStyle);
                        var probe = PdfCore.PdfDocument.Create(frame);
                        var blocks = probe.BuildFlowBlocks(CreateNativeRunningStory(story, section,
                            probeOptions, fontMap, listMarkers, defaults, footer)(context));
                        double? height = PdfCore.PdfWriter.MeasureRunningContentHeight(frame, blocks, width, options.CancellationToken);
                        if (!height.HasValue) return null;
                        maximum = Math.Max(maximum, height.Value);
                    }
                    return maximum;
                }
            }
            options.RunningStoryFits = (section, footer) => !fits.TryGetValue((section, footer), out bool result) || result;
        }
    }
}
