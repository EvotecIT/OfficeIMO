using System;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using A = DocumentFormat.OpenXml.Drawing;
using P = DocumentFormat.OpenXml.Presentation;

namespace OfficeIMO.PowerPoint {
    public sealed partial class PowerPointDocument {
        private const string FooterTotalUri = "urn:officeimo:powerpoint:footer-total:1";

        /// <summary>Adds a real slide-number field and, optionally, a marked ordinary slide-count run.</summary>
        internal static void AddFooterNumber(PowerPointParagraph paragraph, bool includeTotal) {
            paragraph.AddField("1", "slidenum");
            if (!includeTotal) return;
            var run = paragraph.AddRun(" / 1");
            run.Run.RunProperties ??= new A.RunProperties();
            var list = run.Run.RunProperties.GetFirstChild<A.ExtensionList>();
            if (list == null) run.Run.RunProperties.Append(list = new A.ExtensionList());
            var extension = new A.Extension { Uri = FooterTotalUri };
            extension.Append(new OpenXmlUnknownElement("imo", "slideCount", "urn:officeimo:powerpoint:footer:1"));
            list.Append(extension);
        }

        internal static bool IsFooterTotal(A.Run run) => run.RunProperties?.GetFirstChild<A.ExtensionList>()?
            .Elements<A.Extension>().Any(item => item.Uri?.Value == FooterTotalUri) == true;

        /// <summary>Refreshes field caches from the final physical order, including reopened and appended decks.</summary>
        private void RefreshFooterFields() {
            if (LayoutDeck.Footers.Count == 0 && !_slides.Any(slide => slide.NativeFooters.Count != 0)) return;
            for (int index = 0; index < _slides.Count; index++) {
                var root = _slides[index].SlidePart.Slide!;
                var frames = root.CommonSlideData!.ShapeTree!.Descendants<P.Shape>().Where(shape =>
                    shape.NonVisualShapeProperties?.NonVisualDrawingProperties?
                        .GetFirstChild<A.NonVisualDrawingPropertiesExtensionList>()?
                        .Elements<A.NonVisualDrawingPropertiesExtension>().Any(extension => extension.Uri?.Value == FooterSlotUri) == true ||
                    shape.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?.PlaceholderShape?.Type?.Value == P.PlaceholderValues.SlideNumber);
                foreach (var frame in frames) {
                    foreach (var field in frame.Descendants<A.Field>().Where(field => field.Type?.Value == "slidenum")) {
                        field.Text ??= new A.Text();
                        field.Text.Text = (index + 1).ToString(CultureInfo.InvariantCulture);
                    }
                    foreach (var run in frame.Descendants<A.Run>().Where(IsFooterTotal)) {
                        run.Text ??= new A.Text();
                        run.Text.Text = " / " + _slides.Count.ToString(CultureInfo.InvariantCulture);
                    }
                }
            }
        }
    }
}
