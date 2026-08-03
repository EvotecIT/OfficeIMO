using AngleSharp.Dom;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word.Html {
    internal partial class WordToHtmlConverter {
        private IElement CreateSectionElement(IDocument htmlDoc, WordSection section, int index, bool isFirstSection) {
            var element = CreateOutputElement(htmlDoc, "section");
            SetOutputAttribute(htmlDoc, element, "class", "word-section", "SectionMetadata:class");
            SetOutputAttribute(
                htmlDoc,
                element,
                "data-word-section",
                (index + 1).ToString(CultureInfo.InvariantCulture),
                "SectionMetadata:index");
            SetOutputAttribute(
                htmlDoc,
                element,
                "data-page-orientation",
                FormatOrientation(section.PageOrientation),
                "SectionMetadata:orientation");

            var pageSize = section.PageSettings.PageSize;
            if (pageSize != null) {
                SetOutputAttribute(
                    htmlDoc,
                    element,
                    "data-page-size",
                    pageSize.Value.ToString(),
                    "SectionMetadata:page-size");
            }

            var widthTwips = section.PageSettings.Width?.Value;
            var heightTwips = section.PageSettings.Height?.Value;
            if (widthTwips != null) {
                SetOutputAttribute(
                    htmlDoc,
                    element,
                    "data-page-width-twips",
                    widthTwips.Value.ToString(CultureInfo.InvariantCulture),
                    "SectionMetadata:page-width");
            }
            if (heightTwips != null) {
                SetOutputAttribute(
                    htmlDoc,
                    element,
                    "data-page-height-twips",
                    heightTwips.Value.ToString(CultureInfo.InvariantCulture),
                    "SectionMetadata:page-height");
            }

            var top = section.Margins.Top;
            var right = section.Margins.Right?.Value;
            var bottom = section.Margins.Bottom;
            var left = section.Margins.Left?.Value;
            SetTwipsAttribute(htmlDoc, element, "data-margin-top-twips", top, "SectionMetadata:margin-top");
            SetTwipsAttribute(htmlDoc, element, "data-margin-right-twips", right, "SectionMetadata:margin-right");
            SetTwipsAttribute(htmlDoc, element, "data-margin-bottom-twips", bottom, "SectionMetadata:margin-bottom");
            SetTwipsAttribute(htmlDoc, element, "data-margin-left-twips", left, "SectionMetadata:margin-left");

            List<string> styles = new();
            if (widthTwips != null) {
                styles.Add($"width:{FormatTwipsAsPixels(widthTwips.Value)}");
            }
            if (heightTwips != null) {
                styles.Add($"height:{FormatTwipsAsPixels(heightTwips.Value)}");
            }
            if (top != null || right != null || bottom != null || left != null) {
                styles.Add($"padding:{FormatTwipsAsPixels(top ?? 0)} {FormatTwipsAsPixels(right ?? 0)} {FormatTwipsAsPixels(bottom ?? 0)} {FormatTwipsAsPixels(left ?? 0)}");
            }
            if (!isFirstSection) {
                styles.Add("break-before:page");
            }
            SetOutputAttribute(htmlDoc, element, "style", string.Join(";", styles), "SectionMetadata:style");

            return element;
        }

        private static void SetTwipsAttribute(IDocument owner, IElement element, string name, long? value, string source) {
            if (value != null) {
                SetOutputAttribute(owner, element, name, value.Value.ToString(CultureInfo.InvariantCulture), source);
            }
        }

        private static string FormatTwipsAsPixels(long twips) {
            return (twips / 15d).ToString("0.##", CultureInfo.InvariantCulture) + "px";
        }

        private static string FormatOrientation(PageOrientationValues orientation) {
            return orientation == PageOrientationValues.Landscape ? "Landscape" : "Portrait";
        }
    }
}
