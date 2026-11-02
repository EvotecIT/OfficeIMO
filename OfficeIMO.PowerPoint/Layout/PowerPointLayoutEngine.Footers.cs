using System;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointLayoutEngine {
        /// <summary>Renders explicit deck and slide slots using the same grid, styles and native-text binding as slide chrome.</summary>
        internal static bool RenderNativeFooters(PowerPointLayoutContext context, PowerPointLayoutWriter writer,
            IReadOnlyDictionary<OfficeAlignment, PowerPointLayoutFooter>? overrides, bool omitText, bool omitNumber, string? color = null) {
            context.Document.EnsureFooterDefinitions();
            var slide = writer.Slide;
            if (overrides != null) foreach (var item in overrides) slide.NativeFooters[item.Key] = item.Value;
            var effective = new Dictionary<OfficeAlignment, PowerPointLayoutFooter>(context.Deck.Footers);
            foreach (var item in slide.NativeFooters) effective[item.Key] = item.Value;
            if (effective.Count == 0) return false;
            omitText |= context.Deck.NoFooter;
            omitNumber |= !context.Deck.ShowSlideNumber;
            // An explicit footer replaces the implicit deck-title footer and default number frame.
            foreach (var box in slide.TextBoxes.Where(box => !PowerPointDocument.TryReadFooterSlot(box, out _, out _) &&
                         ((Shape)box.Element).NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                             .PlaceholderShape?.Type?.Value is { } type && (type == PlaceholderValues.Footer || type == PlaceholderValues.SlideNumber)).ToArray()) box.Remove();
            var grid = context.Grid;
            double bandHeight = PowerPointLayoutGrid.FooterHeight * context.Theme.Scale;
            var style = context.Style(PowerPointTextRole.Footer);
            color ??= context.Theme.TextMuted;
            if (!omitText) foreach (var item in effective) {
                int column = item.Key == OfficeAlignment.Left ? 1 : item.Key == OfficeAlignment.Center ? 5 : 9;
                int span = item.Key == OfficeAlignment.Left && !effective.ContainsKey(OfficeAlignment.Center) ? 8 : 4;
                var bounds = new PowerPointLayoutRect(grid.X(column), grid.FooterY, grid.W(span), bandHeight);
                var existing = slide.TextBoxes.FirstOrDefault(box => PowerPointDocument.TryReadFooterSlot(box, out var slot, out _) && slot == item.Key);
                if (existing != null) {
                    CheckFits(existing.Paragraphs, bounds.Width);
                    existing.Left = OfficeLength.Points(bounds.X);
                    existing.Top = OfficeLength.Points(bounds.Y);
                    existing.Width = OfficeLength.Points(bounds.Width);
                    existing.Height = OfficeLength.Points(bounds.Height);
                    continue;
                }
                var body = item.Value.TakeBody(omitNumber);
                var paragraphs = body.Paragraphs.Select((paragraph, index) => {
                    var planned = PowerPointLayoutNativeText.Attach(context.Paragraph(paragraph.Text, PowerPointTextRole.Footer, color), body, index);
                    planned.Align = item.Key == OfficeAlignment.Right ? "r" : item.Key == OfficeAlignment.Center ? "ctr" : "l";
                    return planned;
                }).ToArray();
                CheckFits(body.Paragraphs, bounds.Width);
                bool numberOnly = body.Paragraphs.Any(paragraph => paragraph.Paragraph.Elements<A.Field>().Any()) &&
                    body.Paragraphs.All(paragraph => paragraph.Paragraph.ChildElements.All(node => node is A.ParagraphProperties or A.EndParagraphRunProperties ||
                        node is A.Field nativeField && nativeField.Type?.Value == "slidenum" || node is A.Run total && PowerPointDocument.IsFooterTotal(total)));
                uint index = numberOnly ? PowerPointLayoutMasterWriter.SlideNumberIndex
                    : PowerPointLayoutMasterWriter.FooterIndex + (uint)column * 100U;
                string placeholder = "<p:ph type=\"" + (numberOnly ? "sldNum" : "ftr") + "\" idx=\"" + index + "\"/>";
                var frame = writer.Text("Footer " + item.Key, bounds, paragraphs, anchor: "b", placeholder: placeholder, nativeBody: body);
                PowerPointDocument.SetFooterSlot(frame, item.Key, deck: !slide.NativeFooters.ContainsKey(item.Key));
            }
            void CheckFits(IEnumerable<PowerPointParagraph> paragraphs, double width) {
                double height = paragraphs.Sum(paragraph => PowerPointLayoutNativeTextMeasure.Height(context, paragraph,
                    style.Size, width, spaceAfter: 0D, role: PowerPointTextRole.Footer));
                if (height > bandHeight + 0.1D)
                    throw new ArgumentException("The footer content does not fit its footer band. Shorten it or use a smaller font.");
            }
            if (!omitNumber && (omitText || !effective.ContainsKey(OfficeAlignment.Right) && !effective.Values.Any(footer => footer.HasNumber)))
                RenderFooterNumber(context, writer, rightText: false, color, 1D);
            return true;
        }
    }
}
