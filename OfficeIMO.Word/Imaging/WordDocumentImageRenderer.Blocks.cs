using System.Collections.Generic;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private static bool AddBodyElementContent(
            WordDocument document,
            OpenXmlElement element,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers) {
            context.ThrowIfCancellationRequested();
            if (context.TryGetSourceBlock(element, out WordImageSourceBlock sourceBlock)) {
                IReadOnlyList<OfficeDrawingElement> before = CaptureDrawingElements(context.Drawing);
                bool captured = AddBodyElementContentCore(document, element, context, diagnostics, listMarkers);
                CaptureBodyFragment(context, sourceBlock, before);
                return captured;
            }

            return AddBodyElementContentCore(document, element, context, diagnostics, listMarkers);
        }

        private static bool AddBodyElementContentCore(
            WordDocument document,
            OpenXmlElement element,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers) {
            context.ThrowIfCancellationRequested();
            if (element is Paragraph paragraph) {
                return AddBodyParagraphContent(document, paragraph, context, diagnostics, listMarkers);
            }

            if (element is Table table) {
                context.ClearParagraphSpacingState();
                bool added = AddTable(new WordTable(document, table), context, diagnostics, listMarkers);
                context.ClearParagraphSpacingState();
                return added;
            }

            if (element is SdtBlock sdtBlock) {
                return AddBlockContainerContent(document, sdtBlock.SdtContentBlock, context, diagnostics, listMarkers, "body");
            }

            if (element is SdtContentBlock sdtContentBlock) {
                return AddBlockContainerContent(document, sdtContentBlock, context, diagnostics, listMarkers, "body");
            }

            if (element is SectionProperties) {
                return false;
            }

            if (context.IsTargetPage) {
                context.ClearParagraphSpacingState();
                AddDiagnostic(diagnostics, "unsupported-word-body-element", "Skipped a Word body element that is not yet projected through OfficeIMO.Drawing.", element.GetType().Name);
            }

            return false;
        }

        private static bool AddHeaderFooterElementContent(
            WordDocument document,
            OpenXmlElement element,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            string kind) {
            context.ThrowIfCancellationRequested();
            if (element is Paragraph paragraph) {
                if (ShouldSkipParagraphForFinalRevisionView(paragraph)) {
                    return false;
                }

                return AddParagraphContent(document, paragraph, context, diagnostics, listMarkers);
            }

            if (element is Table table) {
                return AddTable(new WordTable(document, table), context, diagnostics, listMarkers);
            }

            if (element is SdtBlock sdtBlock) {
                return AddBlockContainerContent(document, sdtBlock.SdtContentBlock, context, diagnostics, listMarkers, kind);
            }

            if (element is SdtContentBlock sdtContentBlock) {
                return AddBlockContainerContent(document, sdtContentBlock, context, diagnostics, listMarkers, kind);
            }

            AddDiagnostic(diagnostics, "unsupported-word-" + kind + "-element", "Skipped a Word " + kind + " element that is not yet projected through OfficeIMO.Drawing.", element.GetType().Name);
            return false;
        }

        private static bool AddBlockContainerContent(
            WordDocument document,
            OpenXmlElement? container,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            string scope) {
            context.ThrowIfCancellationRequested();
            if (container == null) {
                return false;
            }

            bool addedAny = false;
            List<OpenXmlElement> children = container.ChildElements.ToList();
            for (int index = 0; index < children.Count; index++) {
                context.ThrowIfCancellationRequested();
                OpenXmlElement child = children[index];
                if (scope == "body") {
                    TryAdvanceForKeepWithNext(document, children, index, context, listMarkers);
                    if (context.PastTargetPage || context.StoppedForPagination) {
                        break;
                    }
                }

                bool added = scope == "body"
                    ? AddBodyElementContent(document, child, context, diagnostics, listMarkers)
                    : AddHeaderFooterElementContent(document, child, context, diagnostics, listMarkers, scope);
                addedAny |= added;

                if (context.PastTargetPage || context.StoppedForPagination) {
                    break;
                }

                if (!added && child is Paragraph) {
                    context.ClearParagraphSpacingState();
                    context.Y += ParagraphGapPoints;
                }
            }

            return addedAny;
        }

        private static void TryAdvanceForKeepWithNext(
            WordDocument document,
            IReadOnlyList<OpenXmlElement> elements,
            int currentIndex,
            WordImageFlowContext context,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers) {
            context.ThrowIfCancellationRequested();
            if (currentIndex < 0 ||
                currentIndex >= elements.Count - 1 ||
                elements[currentIndex] is not Paragraph paragraph ||
                !ResolveKeepWithNext(document, paragraph) ||
                context.Y <= context.Top ||
                !context.CanAdvancePageForOverflow) {
                return;
            }

            double keepHeight = EstimateKeepWithNextGroupHeight(
                document,
                elements,
                currentIndex,
                context.ContentWidth,
                listMarkers,
                context.CancellationToken);
            if (keepHeight <= 0D || keepHeight > context.ContentHeight) {
                return;
            }

            if (context.Y + keepHeight > context.ContentBottom) {
                context.AdvanceColumnOrPage();
            }
        }

        private static double EstimateKeepWithNextGroupHeight(
            WordDocument document,
            IReadOnlyList<OpenXmlElement> elements,
            int currentIndex,
            double contentWidth,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            CancellationToken cancellationToken) {
            double height = 0D;
            bool hasFollower = false;
            for (int index = currentIndex; index < elements.Count; index++) {
                cancellationToken.ThrowIfCancellationRequested();
                OpenXmlElement element = elements[index];
                height += EstimateBodyElementHeight(
                    document,
                    element,
                    contentWidth,
                    listMarkers,
                    cancellationToken);
                if (index > currentIndex) {
                    hasFollower = true;
                }

                if (element is not Paragraph paragraph || !ResolveKeepWithNext(document, paragraph)) {
                    break;
                }
            }

            return hasFollower ? height : 0D;
        }

        private static double EstimateBodyElementHeight(
            WordDocument document,
            OpenXmlElement element,
            double contentWidth,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers,
            CancellationToken cancellationToken) {
            var measurementDrawing = new OfficeDrawing(Math.Max(1D, contentWidth), double.MaxValue);
            WordImageFlowContext measurementContext = CreateFlowContext(
                measurementDrawing,
                0D,
                0D,
                Math.Max(1D, contentWidth),
                double.MaxValue,
                "unsupported-word-keep-measurement-overflow",
                "Skipped Word keep-with-next measurement because content does not fit within the measurement frame.",
                cancellationToken: cancellationToken);
            var diagnostics = new List<OfficeImageExportDiagnostic>();
            AddBodyElementContent(document, element, measurementContext, diagnostics, listMarkers);
            return Math.Max(0D, measurementContext.Y);
        }

        private static bool ResolveKeepWithNext(WordDocument document, Paragraph paragraph) {
            KeepNext? direct = paragraph.ParagraphProperties?.GetFirstChild<KeepNext>();
            bool? directValue = ReadOnOff(direct);
            if (directValue.HasValue) {
                return directValue.Value;
            }

            WordParagraph wordParagraph = new WordParagraph(document, paragraph);
            foreach (StyleParagraphProperties properties in EnumerateParagraphStyleProperties(wordParagraph)) {
                bool? styleValue = ReadOnOff(properties.GetFirstChild<KeepNext>());
                if (styleValue.HasValue) {
                    return styleValue.Value;
                }
            }

            return false;
        }

        private static bool ResolveKeepLinesTogether(WordParagraph paragraph) {
            KeepLines? direct = paragraph._paragraphProperties?.GetFirstChild<KeepLines>();
            bool? directValue = ReadOnOff(direct);
            if (directValue.HasValue) {
                return directValue.Value;
            }

            foreach (StyleParagraphProperties properties in EnumerateParagraphStyleProperties(paragraph)) {
                bool? styleValue = ReadOnOff(properties.GetFirstChild<KeepLines>());
                if (styleValue.HasValue) {
                    return styleValue.Value;
                }
            }

            return false;
        }

        private static bool ResolveAvoidWidowAndOrphan(WordParagraph paragraph) {
            WidowControl? direct = paragraph._paragraphProperties?.GetFirstChild<WidowControl>();
            bool? directValue = ReadOnOff(direct);
            if (directValue.HasValue) {
                return directValue.Value;
            }

            foreach (StyleParagraphProperties properties in EnumerateParagraphStyleProperties(paragraph)) {
                bool? styleValue = ReadOnOff(properties.GetFirstChild<WidowControl>());
                if (styleValue.HasValue) {
                    return styleValue.Value;
                }
            }

            WidowControl? defaultValue = paragraph._document?._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles?.DocDefaults?
                .GetFirstChild<ParagraphPropertiesDefault>()?
                .GetFirstChild<ParagraphPropertiesBaseStyle>()?
                .GetFirstChild<WidowControl>();
            bool? resolvedDefault = ReadOnOff(defaultValue);
            return resolvedDefault ?? true;
        }

        private static bool ResolvePageBreakBefore(WordDocument document, Paragraph paragraph) {
            PageBreakBefore? direct = paragraph.ParagraphProperties?.GetFirstChild<PageBreakBefore>();
            bool? directValue = ReadOnOff(direct);
            if (directValue.HasValue) {
                return directValue.Value;
            }

            WordParagraph wordParagraph = new WordParagraph(document, paragraph);
            foreach (StyleParagraphProperties properties in EnumerateParagraphStyleProperties(wordParagraph)) {
                bool? styleValue = ReadOnOff(properties.GetFirstChild<PageBreakBefore>());
                if (styleValue.HasValue) {
                    return styleValue.Value;
                }
            }

            return false;
        }

        private static void AdvanceForPageBreakBefore(WordImageFlowContext context) {
            if (!context.IsAtPageFrameStart) {
                context.AdvancePage();
            }
        }

        private static bool? ReadOnOff(OnOffType? value) {
            if (value == null) {
                return null;
            }

            return value.Val?.Value ?? true;
        }

        private static bool AddBodyParagraphContent(
            WordDocument document,
            Paragraph paragraph,
            WordImageFlowContext context,
            List<OfficeImageExportDiagnostic> diagnostics,
            IReadOnlyDictionary<WordParagraph, (int Level, string Marker)> listMarkers) {
            context.ThrowIfCancellationRequested();
            if (ShouldSkipParagraphForFinalRevisionView(paragraph)) {
                context.ClearParagraphSpacingState();
                return false;
            }

            if (ResolvePageBreakBefore(document, paragraph)) {
                AdvanceForPageBreakBefore(context);
            }

            bool added = AddParagraphContent(document, paragraph, context, diagnostics, listMarkers);
            SectionProperties? sectionProperties = paragraph.ParagraphProperties?.SectionProperties;
            if (IsNextColumnSectionBreak(sectionProperties)) {
                context.AdvanceColumnOrPage();
                return true;
            }

            if (StartsNewPage(sectionProperties)) {
                context.AdvancePages(CountSectionBreakPageAdvance(context.PageIndex, sectionProperties));
                return true;
            }

            return added;
        }

        private static bool ShouldSkipParagraphForFinalRevisionView(Paragraph paragraph) {
            ParagraphMarkRunProperties? paragraphMarkProperties = paragraph.ParagraphProperties?.ParagraphMarkRunProperties;
            if (paragraphMarkProperties == null) {
                return false;
            }

            return paragraphMarkProperties.GetFirstChild<Deleted>() != null ||
                paragraphMarkProperties.GetFirstChild<MoveFrom>() != null;
        }
    }
}
