using AngleSharp.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private IElement CreateSectionElement(IDocument htmlDoc, WordSection section, int index, bool isFirstSection) {
            var element = htmlDoc.CreateElement("section");
            element.SetAttribute("class", "word-section");
            element.SetAttribute("data-word-section", (index + 1).ToString(CultureInfo.InvariantCulture));
            element.SetAttribute("data-page-orientation", FormatOrientation(section.PageOrientation));

            var pageSize = section.PageSettings.PageSize;
            if (pageSize != null) {
                element.SetAttribute("data-page-size", pageSize.Value.ToString());
            }

            var widthTwips = section.PageSettings.Width?.Value;
            var heightTwips = section.PageSettings.Height?.Value;
            if (widthTwips != null) {
                element.SetAttribute("data-page-width-twips", widthTwips.Value.ToString(CultureInfo.InvariantCulture));
            }
            if (heightTwips != null) {
                element.SetAttribute("data-page-height-twips", heightTwips.Value.ToString(CultureInfo.InvariantCulture));
            }

            var top = section.Margins.Top;
            var right = section.Margins.Right?.Value;
            var bottom = section.Margins.Bottom;
            var left = section.Margins.Left?.Value;
            SetTwipsAttribute(element, "data-margin-top-twips", top);
            SetTwipsAttribute(element, "data-margin-right-twips", right);
            SetTwipsAttribute(element, "data-margin-bottom-twips", bottom);
            SetTwipsAttribute(element, "data-margin-left-twips", left);

            List<string> styles = new() { "box-sizing:border-box" };
            if (widthTwips != null) {
                styles.Add($"width:{FormatTwipsAsPoints(widthTwips.Value)}");
            }
            if (heightTwips != null) {
                styles.Add($"min-height:{FormatTwipsAsPoints(heightTwips.Value)}");
            }
            if (top != null || right != null || bottom != null || left != null) {
                styles.Add($"padding:{FormatTwipsAsPoints(top ?? 0)} {FormatTwipsAsPoints(right ?? 0)} {FormatTwipsAsPoints(bottom ?? 0)} {FormatTwipsAsPoints(left ?? 0)}");
            }
            if (!isFirstSection) {
                styles.Add("break-before:page");
            }
            element.SetAttribute("style", string.Join(";", styles));

            return element;
        }

        private static void SetTwipsAttribute(IElement element, string name, long? value) {
            if (value != null) {
                element.SetAttribute(name, value.Value.ToString(CultureInfo.InvariantCulture));
            }
        }

        private static string FormatTwipsAsPoints(long twips) {
            return (twips / 20d).ToString("0.##", CultureInfo.InvariantCulture) + "pt";
        }

        private static string FormatOrientation(PageOrientationValues orientation) {
            return orientation == PageOrientationValues.Landscape ? "Landscape" : "Portrait";
        }
    }
}
