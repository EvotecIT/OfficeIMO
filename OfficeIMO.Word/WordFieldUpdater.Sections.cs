using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    internal static partial class WordFieldUpdater {
        private static bool TryEvaluateSectionNumber(
            WordDocument document,
            MutableFieldCandidate candidate,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Skipped;

            if (candidate.LocationKind != WordFieldLocationKind.Body) {
                message = "SECTION fields outside the document body need Word layout context and were left unchanged.";
                return false;
            }

            int? sectionNumber = EstimateSectionForBodyField(document, candidate.AnchorElement);
            if (sectionNumber == null) {
                message = "SECTION field position could not be matched to a body section.";
                return false;
            }

            return TryFormatSectionInteger(sectionNumber.Value, parsed, "Updated from OfficeIMO body section order.", out value, out status, out message);
        }

        private static bool TryEvaluateSectionPages(
            WordDocument document,
            MutableFieldCandidate candidate,
            WordFieldInventory.ParsedFieldInstruction parsed,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            status = WordFieldUpdateStatus.Skipped;

            if (candidate.LocationKind != WordFieldLocationKind.Body) {
                message = "SECTIONPAGES fields outside the document body need Word layout context and were left unchanged.";
                return false;
            }

            int? sectionPages = EstimateSectionPagesForBodyField(document, candidate.AnchorElement);
            if (sectionPages == null) {
                message = "SECTIONPAGES field position could not be matched to a body section.";
                return false;
            }

            return TryFormatSectionInteger(sectionPages.Value, parsed, "Updated from OfficeIMO page-break count for the containing body section.", out value, out status, out message);
        }

        private static bool TryFormatSectionInteger(
            int numericValue,
            WordFieldInventory.ParsedFieldInstruction parsed,
            string successMessage,
            out string? value,
            out WordFieldUpdateStatus status,
            out string message) {
            value = null;
            WordFieldFormat? format = GetLastMeaningfulFormat(parsed.FormatSwitches);

            if (!string.IsNullOrWhiteSpace(parsed.NumericPictureSwitch) && format != null) {
                status = WordFieldUpdateStatus.Unsupported;
                message = "SECTION and SECTIONPAGES fields cannot combine numeric picture and general format switches for deterministic refresh.";
                return false;
            }

            if (!string.IsNullOrWhiteSpace(parsed.NumericPictureSwitch)) {
                return TryFormatNumericField(numericValue, parsed.NumericPictureSwitch, successMessage, out value, out status, out message);
            }

            switch (format) {
                case null:
                    value = numericValue.ToString(System.Globalization.CultureInfo.InvariantCulture);
                    status = WordFieldUpdateStatus.Updated;
                    message = successMessage;
                    return true;
                case WordFieldFormat.Arabic:
                case WordFieldFormat.Roman:
                case WordFieldFormat.roman:
                case WordFieldFormat.Ordinal:
                case WordFieldFormat.Alphabetical:
                case WordFieldFormat.ALPHABETICAL:
                case WordFieldFormat.Hex:
                case WordFieldFormat.CardText:
                case WordFieldFormat.OrdText:
                case WordFieldFormat.DollarText:
                    value = FormatSequenceValue(numericValue, new[] { format.Value });
                    status = WordFieldUpdateStatus.Updated;
                    message = $"{successMessage} General numeric format switch was applied.";
                    return true;
                default:
                    status = WordFieldUpdateStatus.Unsupported;
                    message = $"SECTION and SECTIONPAGES format switch {format.Value} is not supported for deterministic refresh.";
                    return false;
            }
        }

        private static int? EstimateSectionForBodyField(WordDocument document, OpenXmlElement anchorElement) {
            SectionPageEstimate? estimate = EstimateSectionPageForBodyField(document, anchorElement);
            return estimate?.SectionNumber;
        }

        private static int? EstimateSectionPagesForBodyField(WordDocument document, OpenXmlElement anchorElement) {
            SectionPageEstimate? estimate = EstimateSectionPageForBodyField(document, anchorElement);
            return estimate?.SectionPages;
        }

        private static SectionPageEstimate? EstimateSectionPageForBodyField(WordDocument document, OpenXmlElement anchorElement) {
            Body? body = document._wordprocessingDocument.MainDocumentPart?.Document?.Body;
            if (body == null) {
                return null;
            }

            OpenXmlElement? targetBodyChild = GetBodyChildForField(body, anchorElement);
            if (targetBodyChild == null) {
                return null;
            }

            var pagesBySection = new Dictionary<int, int>();
            int currentSection = 1;
            int targetSection = 0;

            foreach (OpenXmlElement child in body.ChildElements) {
                if (child is SectionProperties) {
                    continue;
                }

                if (!pagesBySection.ContainsKey(currentSection)) {
                    pagesBySection[currentSection] = 1;
                }

                pagesBySection[currentSection] += CountExplicitPageAdvances(child);

                if (ReferenceEquals(child, targetBodyChild)) {
                    targetSection = currentSection;
                }

                if (HasSectionBoundary(child)) {
                    currentSection++;
                }
            }

            if (targetSection == 0 || !pagesBySection.TryGetValue(targetSection, out int sectionPages)) {
                return null;
            }

            return new SectionPageEstimate(targetSection, sectionPages);
        }

        private static OpenXmlElement? GetBodyChildForField(Body body, OpenXmlElement anchorElement) {
            OpenXmlElement current = anchorElement;
            while (current.Parent != null && !ReferenceEquals(current.Parent, body)) {
                current = current.Parent;
            }

            return ReferenceEquals(current.Parent, body) ? current : null;
        }

        private static int CountExplicitPageAdvances(OpenXmlElement element) {
            int count = 0;
            foreach (Paragraph paragraph in EnumerateParagraphs(element)) {
                if (paragraph.ParagraphProperties?.PageBreakBefore != null) {
                    count++;
                }

                count += paragraph
                    .Descendants<Break>()
                    .Count(documentBreak => documentBreak.Type?.Value == BreakValues.Page);
            }

            return count;
        }

        private static IEnumerable<Paragraph> EnumerateParagraphs(OpenXmlElement element) {
            if (element is Paragraph paragraph) {
                yield return paragraph;
            }

            foreach (Paragraph descendant in element.Descendants<Paragraph>()) {
                yield return descendant;
            }
        }

        private static bool HasSectionBoundary(OpenXmlElement element) =>
            element is Paragraph paragraph && paragraph.ParagraphProperties?.SectionProperties != null;

        private readonly struct SectionPageEstimate {
            internal SectionPageEstimate(int sectionNumber, int sectionPages) {
                SectionNumber = sectionNumber;
                SectionPages = sectionPages;
            }

            internal int SectionNumber { get; }

            internal int SectionPages { get; }
        }
    }
}
