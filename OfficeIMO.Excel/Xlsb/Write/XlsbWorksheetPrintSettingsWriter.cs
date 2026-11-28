using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Excel.Xlsb.Projection;
using OfficeIMO.Excel.Xlsb.Biff12;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Validates and writes standard worksheet print options and page margins.</summary>
    internal static class XlsbWorksheetPrintSettingsWriter {
        private const int BrtMargins = 476;
        private const int BrtPrintOptions = 477;
        private const int BrtPageSetup = 478;
        private const int BrtBeginHeaderFooter = 479;
        private const int BrtEndHeaderFooter = 480;

        private static readonly string[] HeaderNames = {
            "oddHeader", "oddFooter", "evenHeader", "evenFooter", "firstHeader", "firstFooter"
        };

        internal static void Write(Stream output, OfficeXmlNode worksheet, string sheetName) {
            ArgumentNullException.ThrowIfNull(output);
            OfficeXmlNode options = XlsbWorksheetMarkup.Child(worksheet, XlsbWorksheetPrintSettingsProjector.Options);
            OfficeXmlNode margins = XlsbWorksheetMarkup.Child(worksheet, XlsbWorksheetPrintSettingsProjector.Margins);
            OfficeXmlNode pageSetup = XlsbWorksheetMarkup.Child(worksheet, XlsbWorksheetPrintSettingsProjector.Setup);
            OfficeXmlNode headerFooter = XlsbWorksheetMarkup.Child(worksheet, XlsbWorksheetPrintSettingsProjector.Headers);
            if (options.Exists) XlsbRecordWriter.Write(output, BrtPrintOptions, CreatePrintOptionsPayload(options, sheetName));
            if (margins.Exists) XlsbRecordWriter.Write(output, BrtMargins, CreateMarginsPayload(margins, sheetName));
            if (pageSetup.Exists) XlsbRecordWriter.Write(output, BrtPageSetup, CreatePageSetupPayload(pageSetup, sheetName));
            if (headerFooter.Exists) {
                XlsbRecordWriter.Write(output, BrtBeginHeaderFooter, CreateHeaderFooterPayload(headerFooter, sheetName));
                XlsbRecordWriter.Write(output, BrtEndHeaderFooter);
            }
        }

        private static byte[] CreatePrintOptionsPayload(OfficeXmlNode options, string sheetName) {
            if (options.FirstChild.Exists) {
                throw new NotSupportedException($"Native XLSB generation does not support child content in print options on worksheet '{sheetName}'.");
            }
            EnsureOnlyAttributes(options, sheetName, "horizontalCentered", "verticalCentered", "headings", "gridLines", "gridLinesSet");
            ushort flags = (ushort)((XlsbWorkbookMarkup.ReadBoolean(options, "horizontalCentered") == true ? 0x0001 : 0)
                | (XlsbWorkbookMarkup.ReadBoolean(options, "verticalCentered") == true ? 0x0002 : 0)
                | (XlsbWorkbookMarkup.ReadBoolean(options, "headings") == true ? 0x0004 : 0)
                | (XlsbWorkbookMarkup.ReadBoolean(options, "gridLines") == true ? 0x0008 : 0));
            return new[] { (byte)flags, (byte)(flags >> 8) };
        }

        private static byte[] CreateMarginsPayload(OfficeXmlNode margins, string sheetName) {
            if (margins.FirstChild.Exists) {
                throw new NotSupportedException($"Native XLSB generation does not support child content in page margins on worksheet '{sheetName}'.");
            }
            EnsureOnlyAttributes(margins, sheetName, "left", "right", "top", "bottom", "header", "footer");
            double[] values = {
                RequireMargin(margins, "left", sheetName),
                RequireMargin(margins, "right", sheetName),
                RequireMargin(margins, "top", sheetName),
                RequireMargin(margins, "bottom", sheetName),
                RequireMargin(margins, "header", sheetName),
                RequireMargin(margins, "footer", sheetName)
            };
            using var output = new MemoryStream(48);
            foreach (double value in values) {
                XlsbPayloadWriter.WriteDouble(output, value);
            }
            return output.ToArray();
        }

        private static byte[] CreatePageSetupPayload(OfficeXmlNode pageSetup, string sheetName) {
            if (pageSetup.FirstChild.Exists) {
                throw new NotSupportedException($"Native XLSB generation does not support child content in page setup on worksheet '{sheetName}'.");
            }
            if (!string.IsNullOrEmpty(pageSetup.Attribute(OfficeXmlNamespaces.Relationships, "id"))) {
                throw new NotSupportedException($"Native XLSB generation does not yet support printer-settings relationships on worksheet '{sheetName}'.");
            }
            EnsureOnlyAttributes(
                pageSetup,
                sheetName,
                "paperSize",
                "scale",
                "firstPageNumber",
                "fitToWidth",
                "fitToHeight",
                "pageOrder",
                "orientation",
                "usePrinterDefaults",
                "blackAndWhite",
                "draft",
                "cellComments",
                "useFirstPageNumber",
                "errors",
                "horizontalDpi",
                "verticalDpi",
                "copies");

            uint scale = XlsbWorkbookMarkup.ReadUInt32(pageSetup, "scale") ?? 0U;
            uint copies = XlsbWorkbookMarkup.ReadUInt32(pageSetup, "copies") ?? 0U;
            uint firstPageNumber = XlsbWorkbookMarkup.ReadUInt32(pageSetup, "firstPageNumber") ?? 0U;
            uint fitToWidth = XlsbWorkbookMarkup.ReadUInt32(pageSetup, "fitToWidth") ?? 0U;
            uint fitToHeight = XlsbWorkbookMarkup.ReadUInt32(pageSetup, "fitToHeight") ?? 0U;
            uint paperSize = XlsbWorkbookMarkup.ReadUInt32(pageSetup, "paperSize") ?? 0U;
            if (paperSize >= int.MaxValue
                || (paperSize >= 119U && paperSize <= 256U)
                || (scale != 0U && (scale < 10U || scale > 400U))
                || copies > 32_767U
                || firstPageNumber > 32_767U
                || fitToWidth > 32_767U
                || fitToHeight > 32_767U) {
                throw new NotSupportedException($"Native XLSB generation cannot encode one or more out-of-range page-setup values on worksheet '{sheetName}'.");
            }

            string pageOrder = pageSetup.Attribute("pageOrder") ?? "downThenOver";
            string orientation = pageSetup.Attribute("orientation") ?? "default";
            string comments = pageSetup.Attribute("cellComments") ?? "none";
            string errors = pageSetup.Attribute("errors") ?? "displayed";
            if (pageOrder is not ("downThenOver" or "overThenDown")
                || orientation is not ("default" or "portrait" or "landscape")
                || comments is not ("none" or "asDisplayed" or "atEnd")
                || errors is not ("displayed" or "blank" or "dash" or "NA")) {
                throw new NotSupportedException($"Native XLSB generation cannot encode one or more unknown page-setup enum values on worksheet '{sheetName}'.");
            }
            ushort flags = 0;
            if (pageOrder == "overThenDown") flags |= 0x0001;
            if (orientation == "landscape") flags |= 0x0002;
            if (XlsbWorkbookMarkup.ReadBoolean(pageSetup, "blackAndWhite") == true) flags |= 0x0008;
            if (XlsbWorkbookMarkup.ReadBoolean(pageSetup, "draft") == true) flags |= 0x0010;
            if (comments != "none") flags |= 0x0020;
            if (orientation == "default" || XlsbWorkbookMarkup.ReadBoolean(pageSetup, "usePrinterDefaults") == true) flags |= 0x0040;
            if (XlsbWorkbookMarkup.ReadBoolean(pageSetup, "useFirstPageNumber") == true) flags |= 0x0080;
            if (comments == "atEnd") flags |= 0x0100;
            if (errors == "blank") flags |= 0x0200;
            else if (errors == "dash") flags |= 0x0400;
            else if (errors == "NA") flags |= 0x0600;

            using var output = new MemoryStream(38);
            XlsbPayloadWriter.WriteUInt32(output, paperSize);
            XlsbPayloadWriter.WriteUInt32(output, scale);
            XlsbPayloadWriter.WriteUInt32(output, XlsbWorkbookMarkup.ReadUInt32(pageSetup, "horizontalDpi") ?? 0U);
            XlsbPayloadWriter.WriteUInt32(output, XlsbWorkbookMarkup.ReadUInt32(pageSetup, "verticalDpi") ?? 0U);
            XlsbPayloadWriter.WriteUInt32(output, copies);
            XlsbPayloadWriter.WriteUInt32(output, firstPageNumber);
            XlsbPayloadWriter.WriteUInt32(output, fitToWidth);
            XlsbPayloadWriter.WriteUInt32(output, fitToHeight);
            XlsbPayloadWriter.WriteUInt16(output, flags);
            XlsbPayloadWriter.WriteUInt32(output, uint.MaxValue);
            return output.ToArray();
        }

        private static byte[] CreateHeaderFooterPayload(OfficeXmlNode headerFooter, string sheetName) {
            EnsureOnlyAttributes(headerFooter, sheetName, "differentOddEven", "differentFirst", "scaleWithDoc", "alignWithMargins");
            var seen = new HashSet<string>(StringComparer.Ordinal);
            foreach (OfficeXmlNode child in headerFooter.Children()) {
                if (child.NamespaceId != OfficeXmlNamespaces.SpreadsheetML || Array.IndexOf(HeaderNames, child.LocalName) < 0 || child.FirstChild.Exists) {
                    throw new NotSupportedException($"Native XLSB generation does not support header/footer content '{child.LocalName}' on worksheet '{sheetName}'.");
                }
                if (!seen.Add(child.LocalName)) {
                    throw new NotSupportedException($"Native XLSB generation requires at most one '{child.LocalName}' element on worksheet '{sheetName}'.");
                }
                EnsureOnlyAttributes(child, sheetName);
            }

            ushort flags = (ushort)((XlsbWorkbookMarkup.ReadBoolean(headerFooter, "differentOddEven") == true ? 0x0001 : 0)
                | (XlsbWorkbookMarkup.ReadBoolean(headerFooter, "differentFirst") == true ? 0x0002 : 0)
                | (XlsbWorkbookMarkup.ReadBoolean(headerFooter, "scaleWithDoc") != false ? 0x0004 : 0)
                | (XlsbWorkbookMarkup.ReadBoolean(headerFooter, "alignWithMargins") != false ? 0x0008 : 0));
            using var output = new MemoryStream();
            XlsbPayloadWriter.WriteUInt16(output, flags);
            foreach (string name in HeaderNames) {
                WriteNullableWideString(output, headerFooter.ChildText(OfficeXmlNamespaces.SpreadsheetML, name), sheetName);
            }
            return output.ToArray();
        }

        private static double RequireMargin(OfficeXmlNode margins, string detail, string sheetName) {
            string? attribute = margins.Attribute(detail);
            if (attribute == null) throw MarginRequired(detail, sheetName);
            double value = System.Xml.XmlConvert.ToDouble(attribute);
            if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D || value >= 49D) {
                throw MarginRequired(detail, sheetName);
            }
            return value;
        }

        private static NotSupportedException MarginRequired(string detail, string sheetName) =>
            new($"Native XLSB generation requires a {detail} page margin from 0 up to 49 inches on worksheet '{sheetName}'.");

        private static void WriteNullableWideString(Stream output, string? value, string sheetName) {
            if (value == null) {
                XlsbPayloadWriter.WriteUInt32(output, uint.MaxValue);
                return;
            }
            if (value.Length > 255) {
                throw new NotSupportedException($"Native XLSB generation limits each header/footer string to 255 characters on worksheet '{sheetName}'.");
            }
            XlsbPayloadWriter.WriteUInt32(output, checked((uint)value.Length));
            byte[] bytes = Encoding.Unicode.GetBytes(value);
            output.Write(bytes, 0, bytes.Length);
        }

        private static void EnsureOnlyAttributes(OfficeXmlNode element, string sheetName, params string[] allowedNames) {
            string? unsupported = XlsbWorkbookMarkup.UnsupportedAttribute(element, allowedNames);
            if (unsupported != null) {
                throw new NotSupportedException($"Native XLSB generation does not yet support print-setting attribute '{unsupported}' on worksheet '{sheetName}'.");
            }
        }
    }
}
