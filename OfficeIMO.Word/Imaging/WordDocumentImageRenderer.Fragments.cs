using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml;
using OfficeIMO.Drawing;

namespace OfficeIMO.Word {
    internal static partial class WordDocumentImageRenderer {
        private static IReadOnlyDictionary<OpenXmlElement, WordImageSourceBlock> BuildImageSourceBlocks(
            WordDocument document,
            CancellationToken cancellationToken = default) {
            var result = new Dictionary<OpenXmlElement, WordImageSourceBlock>(
                OpenXmlElementReferenceComparer.Instance);

            for (int sectionIndex = 0; sectionIndex < document.Sections.Count; sectionIndex++) {
                cancellationToken.ThrowIfCancellationRequested();
                var seen = new HashSet<OpenXmlElement>(OpenXmlElementReferenceComparer.Instance);
                int blockIndex = 0;
                foreach (var element in document.Sections[sectionIndex].Elements) {
                    cancellationToken.ThrowIfCancellationRequested();
                    OpenXmlElement? source = null;
                    string? kind = null;
                    if (element is WordParagraph paragraph) {
                        source = paragraph._paragraph;
                        kind = "paragraph";
                    } else if (element is WordTable table) {
                        source = table._table;
                        kind = "table";
                    }

                    if (source == null || !seen.Add(source)) {
                        continue;
                    }

                    result[source] = new WordImageSourceBlock(sectionIndex, blockIndex, kind!);
                    blockIndex++;
                }
            }

            return result;
        }

        private static IReadOnlyList<OfficeDrawingElement> CaptureDrawingElements(OfficeDrawing drawing) =>
            drawing.Elements.ToArray();

        private static void CaptureBodyFragment(
            WordImageFlowContext context,
            WordImageSourceBlock source,
            IReadOnlyList<OfficeDrawingElement> before) {
            IReadOnlyList<OfficeDrawingElement> current = context.Drawing.Elements;
            if (current.Count <= before.Count) {
                return;
            }

            var existing = new HashSet<OfficeDrawingElement>(before, DrawingElementReferenceComparer.Instance);
            var text = new StringBuilder();
            bool hasBounds = false;
            double left = double.MaxValue;
            double top = double.MaxValue;
            double right = double.MinValue;
            double bottom = double.MinValue;

            for (int index = 0; index < current.Count; index++) {
                OfficeDrawingElement element = current[index];
                if (existing.Contains(element)) {
                    continue;
                }

                AppendVisibleText(element, text);
                if (TryGetBounds(element, out double elementLeft, out double elementTop, out double elementRight, out double elementBottom)) {
                    left = Math.Min(left, elementLeft);
                    top = Math.Min(top, elementTop);
                    right = Math.Max(right, elementRight);
                    bottom = Math.Max(bottom, elementBottom);
                    hasBounds = true;
                }
            }

            WordDocumentVisualRegion? region = hasBounds
                ? new WordDocumentVisualRegion(left, top, Math.Max(0D, right - left), Math.Max(0D, bottom - top))
                : null;
            context.AddFragment(new WordDocumentVisualFragment(
                source.SectionIndex,
                source.BlockIndex,
                source.Kind,
                text.ToString().Trim(),
                region));
        }

        private static void AppendVisibleText(OfficeDrawingElement element, StringBuilder text) {
            string? value = element switch {
                OfficeDrawingText drawingText => drawingText.Text,
                OfficeDrawingRichText richText => richText.PlainText,
                _ => null
            };
            if (!string.IsNullOrWhiteSpace(value)) {
                if (text.Length > 0) {
                    text.AppendLine();
                }
                text.Append(value);
            }

            if (element is OfficeDrawingGroup group) {
                foreach (OfficeDrawingElement child in group.Drawing.Elements) {
                    AppendVisibleText(child, text);
                }
            } else if (element is OfficeDrawingEffectGroup effectGroup) {
                foreach (OfficeDrawingElement child in effectGroup.Drawing.Elements) {
                    AppendVisibleText(child, text);
                }
            }
        }

        private static bool TryGetBounds(
            OfficeDrawingElement element,
            out double left,
            out double top,
            out double right,
            out double bottom) {
            switch (element) {
                case OfficeDrawingText text:
                    left = text.X;
                    top = text.Y;
                    right = text.X + text.Width;
                    bottom = text.Y + text.Height;
                    return true;
                case OfficeDrawingRichText richText:
                    left = richText.X;
                    top = richText.Y;
                    right = richText.X + richText.Width;
                    bottom = richText.Y + richText.Height;
                    return true;
                case OfficeDrawingShape shape:
                    left = shape.X;
                    top = shape.Y;
                    right = shape.X + shape.Shape.Width;
                    bottom = shape.Y + shape.Shape.Height;
                    return true;
                case OfficeDrawingImage image:
                    (left, top, right, bottom) = image.Projection.GetDestinationBounds();
                    return true;
                case OfficeDrawingGroup group:
                    left = group.X;
                    top = group.Y;
                    right = group.X + group.ClipPath.Width;
                    bottom = group.Y + group.ClipPath.Height;
                    return true;
                default:
                    left = 0D;
                    top = 0D;
                    right = 0D;
                    bottom = 0D;
                    return false;
            }
        }

        private readonly struct WordImageSourceBlock {
            internal WordImageSourceBlock(int sectionIndex, int blockIndex, string kind) {
                SectionIndex = sectionIndex;
                BlockIndex = blockIndex;
                Kind = kind;
            }

            internal int SectionIndex { get; }

            internal int BlockIndex { get; }

            internal string Kind { get; }
        }

        private sealed class OpenXmlElementReferenceComparer : IEqualityComparer<OpenXmlElement> {
            internal static OpenXmlElementReferenceComparer Instance { get; } = new OpenXmlElementReferenceComparer();

            public bool Equals(OpenXmlElement? x, OpenXmlElement? y) => ReferenceEquals(x, y);

            public int GetHashCode(OpenXmlElement obj) =>
                System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
        }

        private sealed class DrawingElementReferenceComparer : IEqualityComparer<OfficeDrawingElement> {
            internal static DrawingElementReferenceComparer Instance { get; } = new DrawingElementReferenceComparer();

            public bool Equals(OfficeDrawingElement? x, OfficeDrawingElement? y) => ReferenceEquals(x, y);

            public int GetHashCode(OfficeDrawingElement obj) =>
                System.Runtime.CompilerServices.RuntimeHelpers.GetHashCode(obj);
        }
    }
}
