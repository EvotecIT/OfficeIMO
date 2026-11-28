using OfficeIMO.OpenXml.Internal;

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
                case WordFieldFormat.UpperRoman:
                case WordFieldFormat.LowerRoman:
                case WordFieldFormat.Ordinal:
                case WordFieldFormat.LowerLetter:
                case WordFieldFormat.UpperLetter:
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

        private static int? EstimateSectionForBodyField(WordDocument document, OfficeXmlNode anchorElement) {
            SectionPageEstimate? estimate = EstimateSectionPageForBodyField(document, anchorElement);
            return estimate?.SectionNumber;
        }

        private static int? EstimateSectionPagesForBodyField(WordDocument document, OfficeXmlNode anchorElement) {
            SectionPageEstimate? estimate = EstimateSectionPageForBodyField(document, anchorElement);
            return estimate?.SectionPages;
        }

        private static SectionPageEstimate? EstimateSectionPageForBodyField(WordDocument document, OfficeXmlNode anchorElement) {
            OfficeXmlNode body = WordFieldInventory.BodyNode(document);
            if (body.IsNone) return null;

            OfficeXmlNode targetBodyChild = GetBodyChildForField(body, anchorElement);
            if (targetBodyChild.IsNone) return null;

            var pagesBySection = new Dictionary<int, int>();
            int currentSection = 1;
            int targetSection = 0;

            foreach (OfficeXmlNode child in body.Children()) {
                if (child.Is(FieldNamespace, "sectPr")) {
                    continue;
                }

                if (!pagesBySection.ContainsKey(currentSection)) {
                    pagesBySection[currentSection] = 1;
                }

                pagesBySection[currentSection] += CountExplicitPageAdvances(child);

                if (child == targetBodyChild) {
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

        private static OfficeXmlNode GetBodyChildForField(OfficeXmlNode body, OfficeXmlNode anchorElement) {
            for (OfficeXmlNode current = anchorElement; current.Exists; current = current.Parent) {
                if (current.Parent == body) return current;
            }
            return default;
        }

        private static int CountExplicitPageAdvances(OfficeXmlNode element) {
            int count = element.Is(FieldNamespace, "p") ? CountParagraphPageBreaks(element) : 0;
            foreach (OfficeXmlNode node in WordFieldInventory.EnumerateFieldContent(element)) {
                if (node.Is(FieldNamespace, "p")) count += CountParagraphPageBreaks(node);
            }
            return count;
        }

        private static bool HasSectionBoundary(OfficeXmlNode element) => element.Is(FieldNamespace, "p") &&
            element.Child(FieldNamespace, "pPr").Child(FieldNamespace, "sectPr").Exists;

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
