using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Writes the normal workbook-window records referenced by worksheet views.</summary>
    internal static class XlsbWorkbookViewWriter {
        private const int BrtBeginBookViews = 135;
        private const int BrtEndBookViews = 136;
        private const int BrtBookView = 158;

        internal static void Write(Stream output, BookViews? views, int sheetCount) {
            if (output == null) throw new ArgumentNullException(nameof(output));
            if (views == null) return;
            if (sheetCount <= 0) throw new ArgumentOutOfRangeException(nameof(sheetCount));
            EnsureOnlyAttributes(views);
            WorkbookView[] workbookViews = views.Elements<WorkbookView>().ToArray();
            if (workbookViews.Length != views.ChildElements.Count || workbookViews.Length == 0) {
                throw new NotSupportedException("Native XLSB generation requires workbook views to contain one or more normal workbookView elements.");
            }

            XlsbRecordWriter.Write(output, BrtBeginBookViews);
            foreach (WorkbookView view in workbookViews) {
                XlsbRecordWriter.Write(output, BrtBookView, CreatePayload(view, sheetCount));
            }
            XlsbRecordWriter.Write(output, BrtEndBookViews);
        }

        internal static void Validate(BookViews? views, int sheetCount) {
            if (views == null) return;
            using var sink = new MemoryStream();
            Write(sink, views, sheetCount);
        }

        private static byte[] CreatePayload(WorkbookView view, int sheetCount) {
            EnsureOnlyAttributes(view,
                "visibility", "minimized", "showHorizontalScroll", "showVerticalScroll", "showSheetTabs",
                "xWindow", "yWindow", "windowWidth", "windowHeight", "tabRatio", "firstSheet", "activeTab",
                "autoFilterDateGrouping");
            if (view.HasChildren) {
                throw new NotSupportedException("Native XLSB generation does not support child content in workbookView elements.");
            }

            uint activeTab = CoerceSheetIndex(view.ActiveTab?.Value, sheetCount, 0U);
            uint firstSheet = CoerceSheetIndex(view.FirstSheet?.Value, sheetCount, activeTab);
            uint width = view.WindowWidth?.Value ?? 28_800U;
            uint height = view.WindowHeight?.Value ?? 17_640U;
            uint tabRatio = view.TabRatio?.Value ?? 600U;
            if (width == 0U || height == 0U || tabRatio > 1000U) {
                throw new NotSupportedException("Native XLSB generation requires positive workbook window dimensions and a tabRatio from 0 through 1000.");
            }

            VisibilityValues? visibility = view.Visibility?.Value;
            byte flags = 0;
            if (visibility == VisibilityValues.Hidden || visibility == VisibilityValues.VeryHidden) flags |= 0x01;
            if (view.Minimized?.Value == true) flags |= 0x02;
            if (visibility == VisibilityValues.VeryHidden) flags |= 0x04;
            if (view.ShowHorizontalScroll?.Value ?? true) flags |= 0x08;
            if (view.ShowVerticalScroll?.Value ?? true) flags |= 0x10;
            if (view.ShowSheetTabs?.Value ?? true) flags |= 0x20;
            if (view.AutoFilterDateGrouping?.Value ?? true) flags |= 0x40;

            using var payload = new MemoryStream(29);
            WriteInt32(payload, view.XWindow?.Value ?? 0);
            WriteInt32(payload, view.YWindow?.Value ?? 0);
            WriteUInt32(payload, width);
            WriteUInt32(payload, height);
            WriteUInt32(payload, tabRatio);
            WriteUInt32(payload, firstSheet);
            WriteUInt32(payload, activeTab);
            payload.WriteByte(flags);
            return payload.ToArray();
        }

        private static uint CoerceSheetIndex(uint? index, int sheetCount, uint fallback) =>
            index.HasValue && index.Value < sheetCount ? index.Value : fallback;

        private static void EnsureOnlyAttributes(OpenXmlElement element, params string[] allowedNames) {
            var allowed = new HashSet<string>(allowedNames, StringComparer.Ordinal);
            OpenXmlAttribute? unsupported = element.GetAttributes()
                .Cast<OpenXmlAttribute?>()
                .FirstOrDefault(attribute => attribute.HasValue
                    && !string.Equals(attribute.Value.NamespaceUri, "http://www.w3.org/2000/xmlns/", StringComparison.Ordinal)
                    && !allowed.Contains(attribute.Value.LocalName));
            if (unsupported.HasValue) {
                throw new NotSupportedException($"Native XLSB generation does not yet support workbook-view attribute '{unsupported.Value.LocalName}'.");
            }
        }

        private static void WriteInt32(Stream stream, int value) => WriteUInt32(stream, unchecked((uint)value));

        private static void WriteUInt32(Stream stream, uint value) {
            stream.WriteByte((byte)value);
            stream.WriteByte((byte)(value >> 8));
            stream.WriteByte((byte)(value >> 16));
            stream.WriteByte((byte)(value >> 24));
        }
    }
}
