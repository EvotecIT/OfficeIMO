using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Projection;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Writes the normal workbook-window records referenced by worksheet views.</summary>
    internal static class XlsbWorkbookViewWriter {
        private const int BrtBeginBookViews = 135;
        private const int BrtEndBookViews = 136;
        private const int BrtBookView = 158;
        internal const string ElementName = "bookViews";

        private static readonly string[] ViewAttributes = {
            "visibility", "minimized", "showHorizontalScroll", "showVerticalScroll", "showSheetTabs",
            "xWindow", "yWindow", "windowWidth", "windowHeight", "tabRatio", "firstSheet", "activeTab",
            "autoFilterDateGrouping"
        };

        internal static void Write(Stream output, OfficeXmlNode views, int sheetCount) {
            ArgumentNullException.ThrowIfNull(output);
            if (views.IsNone) return;
            ArgumentOutOfRangeException.ThrowIfNegativeOrZero(sheetCount);
            EnsureOnlyAttributes(views, Array.Empty<string>());
            if (!views.FirstChild.Exists) ThrowUnsupportedViews();
            foreach (OfficeXmlNode view in views.Children()) {
                if (!view.Is(OfficeXmlNamespaces.SpreadsheetML, "workbookView")) ThrowUnsupportedViews();
            }

            XlsbRecordWriter.Write(output, BrtBeginBookViews);
            foreach (OfficeXmlNode view in views.Children()) {
                XlsbRecordWriter.Write(output, BrtBookView, CreatePayload(view, sheetCount));
            }
            XlsbRecordWriter.Write(output, BrtEndBookViews);
        }

        internal static void Validate(OfficeXmlNode views, int sheetCount) {
            Write(Stream.Null, views, sheetCount);
        }

        private static byte[] CreatePayload(OfficeXmlNode view, int sheetCount) {
            EnsureOnlyAttributes(view, ViewAttributes);
            if (view.FirstChild.Exists) {
                throw new NotSupportedException("Native XLSB generation does not support child content in workbookView elements.");
            }

            uint activeTab = CoerceSheetIndex(XlsbWorkbookMarkup.ReadUInt32(view, "activeTab"), sheetCount, 0U);
            uint firstSheet = CoerceSheetIndex(XlsbWorkbookMarkup.ReadUInt32(view, "firstSheet"), sheetCount, activeTab);
            uint width = XlsbWorkbookMarkup.ReadUInt32(view, "windowWidth") ?? 28_800U;
            uint height = XlsbWorkbookMarkup.ReadUInt32(view, "windowHeight") ?? 17_640U;
            uint tabRatio = XlsbWorkbookMarkup.ReadUInt32(view, "tabRatio") ?? 600U;
            if (width == 0U || height == 0U || tabRatio > 1000U) {
                throw new NotSupportedException("Native XLSB generation requires positive workbook window dimensions and a tabRatio from 0 through 1000.");
            }

            byte flags = view.Attribute("visibility") switch {
                null or "visible" => 0,
                "hidden" => 0x01,
                "veryHidden" => 0x05,
                string value => throw new FormatException($"Invalid workbook view visibility '{value}'.")
            };
            if (XlsbWorkbookMarkup.ReadBoolean(view, "minimized") == true) flags |= 0x02;
            if (XlsbWorkbookMarkup.ReadBoolean(view, "showHorizontalScroll") ?? true) flags |= 0x08;
            if (XlsbWorkbookMarkup.ReadBoolean(view, "showVerticalScroll") ?? true) flags |= 0x10;
            if (XlsbWorkbookMarkup.ReadBoolean(view, "showSheetTabs") ?? true) flags |= 0x20;
            if (XlsbWorkbookMarkup.ReadBoolean(view, "autoFilterDateGrouping") ?? true) flags |= 0x40;

            using var payload = new MemoryStream(29);
            XlsbPayloadWriter.WriteInt32(payload, XlsbWorkbookMarkup.ReadInt32(view, "xWindow") ?? 0);
            XlsbPayloadWriter.WriteInt32(payload, XlsbWorkbookMarkup.ReadInt32(view, "yWindow") ?? 0);
            XlsbPayloadWriter.WriteUInt32(payload, width);
            XlsbPayloadWriter.WriteUInt32(payload, height);
            XlsbPayloadWriter.WriteUInt32(payload, tabRatio);
            XlsbPayloadWriter.WriteUInt32(payload, firstSheet);
            XlsbPayloadWriter.WriteUInt32(payload, activeTab);
            payload.WriteByte(flags);
            return payload.ToArray();
        }

        private static uint CoerceSheetIndex(uint? index, int sheetCount, uint fallback) =>
            index.HasValue && index.Value < sheetCount ? index.Value : fallback;

        private static void EnsureOnlyAttributes(OfficeXmlNode element, string[] allowedNames) {
            string? unsupported = XlsbWorkbookMarkup.UnsupportedAttribute(element, allowedNames);
            if (unsupported != null) {
                throw new NotSupportedException($"Native XLSB generation does not yet support workbook-view attribute '{unsupported}'.");
            }
        }

        private static void ThrowUnsupportedViews() =>
            throw new NotSupportedException("Native XLSB generation requires workbook views to contain one or more normal workbookView elements.");
    }
}
