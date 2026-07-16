using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;

namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Validates and writes standard worksheet print options and page margins.</summary>
    internal static class XlsbWorksheetPrintSettingsWriter {
        private const int BrtMargins = 476;
        private const int BrtPrintOptions = 477;
        private const int BrtPageSetup = 478;
        private const int BrtBeginHeaderFooter = 479;
        private const int BrtEndHeaderFooter = 480;

        internal static void Write(
            Stream output,
            PrintOptions? options,
            PageMargins? margins,
            PageSetup? pageSetup,
            HeaderFooter? headerFooter,
            string sheetName) {
            if (output == null) throw new ArgumentNullException(nameof(output));
            if (options != null) XlsbRecordWriter.Write(output, BrtPrintOptions, CreatePrintOptionsPayload(options, sheetName));
            if (margins != null) XlsbRecordWriter.Write(output, BrtMargins, CreateMarginsPayload(margins, sheetName));
            if (pageSetup != null) XlsbRecordWriter.Write(output, BrtPageSetup, CreatePageSetupPayload(pageSetup, sheetName));
            if (headerFooter != null) {
                XlsbRecordWriter.Write(output, BrtBeginHeaderFooter, CreateHeaderFooterPayload(headerFooter, sheetName));
                XlsbRecordWriter.Write(output, BrtEndHeaderFooter);
            }
        }

        internal static void Validate(
            PrintOptions? options,
            PageMargins? margins,
            PageSetup? pageSetup,
            HeaderFooter? headerFooter,
            string sheetName) {
            if (options != null) CreatePrintOptionsPayload(options, sheetName);
            if (margins != null) CreateMarginsPayload(margins, sheetName);
            if (pageSetup != null) CreatePageSetupPayload(pageSetup, sheetName);
            if (headerFooter != null) CreateHeaderFooterPayload(headerFooter, sheetName);
        }

        private static byte[] CreatePrintOptionsPayload(PrintOptions options, string sheetName) {
            if (options.HasChildren) {
                throw new NotSupportedException($"Native XLSB generation does not support child content in print options on worksheet '{sheetName}'.");
            }
            EnsureOnlyAttributes(options, sheetName, "horizontalCentered", "verticalCentered", "headings", "gridLines", "gridLinesSet");
            ushort flags = (ushort)((options.HorizontalCentered?.Value == true ? 0x0001 : 0)
                | (options.VerticalCentered?.Value == true ? 0x0002 : 0)
                | (options.Headings?.Value == true ? 0x0004 : 0)
                | (options.GridLines?.Value == true ? 0x0008 : 0));
            return new[] { (byte)flags, (byte)(flags >> 8) };
        }

        private static byte[] CreateMarginsPayload(PageMargins margins, string sheetName) {
            if (margins.HasChildren) {
                throw new NotSupportedException($"Native XLSB generation does not support child content in page margins on worksheet '{sheetName}'.");
            }
            EnsureOnlyAttributes(margins, sheetName, "left", "right", "top", "bottom", "header", "footer");
            double[] values = {
                RequireMargin(margins.Left, "left", sheetName),
                RequireMargin(margins.Right, "right", sheetName),
                RequireMargin(margins.Top, "top", sheetName),
                RequireMargin(margins.Bottom, "bottom", sheetName),
                RequireMargin(margins.Header, "header", sheetName),
                RequireMargin(margins.Footer, "footer", sheetName)
            };
            using var output = new MemoryStream(48);
            foreach (double value in values) {
                byte[] bytes = BitConverter.GetBytes(value);
                output.Write(bytes, 0, bytes.Length);
            }
            return output.ToArray();
        }

        private static byte[] CreatePageSetupPayload(PageSetup pageSetup, string sheetName) {
            if (pageSetup.HasChildren) {
                throw new NotSupportedException($"Native XLSB generation does not support child content in page setup on worksheet '{sheetName}'.");
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
                "copies",
                "id");

            if (!string.IsNullOrEmpty(pageSetup.Id?.Value)) {
                throw new NotSupportedException($"Native XLSB generation does not yet support printer-settings relationships on worksheet '{sheetName}'.");
            }

            uint scale = pageSetup.Scale?.Value ?? 0U;
            uint copies = pageSetup.Copies?.Value ?? 0U;
            uint firstPageNumber = pageSetup.FirstPageNumber?.Value ?? 0U;
            uint fitToWidth = pageSetup.FitToWidth?.Value ?? 0U;
            uint fitToHeight = pageSetup.FitToHeight?.Value ?? 0U;
            uint paperSize = pageSetup.PaperSize?.Value ?? 0U;
            if (paperSize >= int.MaxValue
                || (paperSize >= 119U && paperSize <= 256U)
                || (scale != 0U && (scale < 10U || scale > 400U))
                || copies > 32_767U
                || firstPageNumber > 32_767U
                || fitToWidth > 32_767U
                || fitToHeight > 32_767U) {
                throw new NotSupportedException($"Native XLSB generation cannot encode one or more out-of-range page-setup values on worksheet '{sheetName}'.");
            }

            PageOrderValues pageOrder = pageSetup.PageOrder?.Value ?? PageOrderValues.DownThenOver;
            OrientationValues orientation = pageSetup.Orientation?.Value ?? OrientationValues.Default;
            CellCommentsValues comments = pageSetup.CellComments?.Value ?? CellCommentsValues.None;
            PrintErrorValues errors = pageSetup.Errors?.Value ?? PrintErrorValues.Displayed;
            if ((pageOrder != PageOrderValues.DownThenOver && pageOrder != PageOrderValues.OverThenDown)
                || (orientation != OrientationValues.Default && orientation != OrientationValues.Portrait && orientation != OrientationValues.Landscape)
                || (comments != CellCommentsValues.None && comments != CellCommentsValues.AsDisplayed && comments != CellCommentsValues.AtEnd)
                || (errors != PrintErrorValues.Displayed && errors != PrintErrorValues.Blank && errors != PrintErrorValues.Dash && errors != PrintErrorValues.NA)) {
                throw new NotSupportedException($"Native XLSB generation cannot encode one or more unknown page-setup enum values on worksheet '{sheetName}'.");
            }
            ushort flags = 0;
            if (pageOrder == PageOrderValues.OverThenDown) flags |= 0x0001;
            if (orientation == OrientationValues.Landscape) flags |= 0x0002;
            if (pageSetup.BlackAndWhite?.Value == true) flags |= 0x0008;
            if (pageSetup.Draft?.Value == true) flags |= 0x0010;
            if (comments != CellCommentsValues.None) flags |= 0x0020;
            if (orientation == OrientationValues.Default || pageSetup.UsePrinterDefaults?.Value == true) flags |= 0x0040;
            if (pageSetup.UseFirstPageNumber?.Value == true) flags |= 0x0080;
            if (comments == CellCommentsValues.AtEnd) flags |= 0x0100;
            if (errors == PrintErrorValues.Blank) flags |= 0x0200;
            else if (errors == PrintErrorValues.Dash) flags |= 0x0400;
            else if (errors == PrintErrorValues.NA) flags |= 0x0600;

            using var output = new MemoryStream(38);
            WriteUInt32(output, paperSize);
            WriteUInt32(output, scale);
            WriteUInt32(output, pageSetup.HorizontalDpi?.Value ?? 0U);
            WriteUInt32(output, pageSetup.VerticalDpi?.Value ?? 0U);
            WriteUInt32(output, copies);
            WriteUInt32(output, firstPageNumber);
            WriteUInt32(output, fitToWidth);
            WriteUInt32(output, fitToHeight);
            WriteUInt16(output, flags);
            WriteUInt32(output, uint.MaxValue);
            return output.ToArray();
        }

        private static byte[] CreateHeaderFooterPayload(HeaderFooter headerFooter, string sheetName) {
            EnsureOnlyAttributes(headerFooter, sheetName, "differentOddEven", "differentFirst", "scaleWithDoc", "alignWithMargins");
            Type[] allowedTypes = {
                typeof(OddHeader), typeof(OddFooter), typeof(EvenHeader),
                typeof(EvenFooter), typeof(FirstHeader), typeof(FirstFooter)
            };
            foreach (OpenXmlElement child in headerFooter.ChildElements) {
                if (!allowedTypes.Contains(child.GetType())) {
                    throw new NotSupportedException($"Native XLSB generation does not support header/footer content '{child.LocalName}' on worksheet '{sheetName}'.");
                }
                if (headerFooter.ChildElements.Count(candidate => candidate.GetType() == child.GetType()) > 1) {
                    throw new NotSupportedException($"Native XLSB generation requires at most one '{child.LocalName}' element on worksheet '{sheetName}'.");
                }
                EnsureOnlyAttributes(child, sheetName);
            }

            ushort flags = (ushort)((headerFooter.DifferentOddEven?.Value == true ? 0x0001 : 0)
                | (headerFooter.DifferentFirst?.Value == true ? 0x0002 : 0)
                | (headerFooter.ScaleWithDoc?.Value != false ? 0x0004 : 0)
                | (headerFooter.AlignWithMargins?.Value != false ? 0x0008 : 0));
            using var output = new MemoryStream();
            WriteUInt16(output, flags);
            WriteNullableWideString(output, headerFooter.GetFirstChild<OddHeader>()?.Text, sheetName);
            WriteNullableWideString(output, headerFooter.GetFirstChild<OddFooter>()?.Text, sheetName);
            WriteNullableWideString(output, headerFooter.GetFirstChild<EvenHeader>()?.Text, sheetName);
            WriteNullableWideString(output, headerFooter.GetFirstChild<EvenFooter>()?.Text, sheetName);
            WriteNullableWideString(output, headerFooter.GetFirstChild<FirstHeader>()?.Text, sheetName);
            WriteNullableWideString(output, headerFooter.GetFirstChild<FirstFooter>()?.Text, sheetName);
            return output.ToArray();
        }

        private static double RequireMargin(DoubleValue? margin, string detail, string sheetName) {
            if (margin?.Value is not double value || double.IsNaN(value) || double.IsInfinity(value) || value < 0D || value >= 49D) {
                throw new NotSupportedException($"Native XLSB generation requires a {detail} page margin from 0 up to 49 inches on worksheet '{sheetName}'.");
            }
            return value;
        }

        private static void WriteNullableWideString(Stream output, string? value, string sheetName) {
            if (value == null) {
                WriteUInt32(output, uint.MaxValue);
                return;
            }
            if (value.Length > 255) {
                throw new NotSupportedException($"Native XLSB generation limits each header/footer string to 255 characters on worksheet '{sheetName}'.");
            }
            WriteUInt32(output, checked((uint)value.Length));
            byte[] bytes = Encoding.Unicode.GetBytes(value);
            output.Write(bytes, 0, bytes.Length);
        }

        private static void WriteUInt16(Stream output, ushort value) {
            output.WriteByte((byte)value);
            output.WriteByte((byte)(value >> 8));
        }

        private static void WriteUInt32(Stream output, uint value) {
            output.WriteByte((byte)value);
            output.WriteByte((byte)(value >> 8));
            output.WriteByte((byte)(value >> 16));
            output.WriteByte((byte)(value >> 24));
        }

        private static void EnsureOnlyAttributes(OpenXmlElement element, string sheetName, params string[] allowedNames) {
            var allowed = new HashSet<string>(allowedNames, StringComparer.Ordinal);
            OpenXmlAttribute? unsupported = element.GetAttributes()
                .Cast<OpenXmlAttribute?>()
                .FirstOrDefault(attribute => attribute.HasValue
                    && !string.Equals(attribute.Value.NamespaceUri, "http://www.w3.org/2000/xmlns/", StringComparison.Ordinal)
                    && !allowed.Contains(attribute.Value.LocalName));
            if (unsupported.HasValue) {
                throw new NotSupportedException($"Native XLSB generation does not yet support print-setting attribute '{unsupported.Value.LocalName}' on worksheet '{sheetName}'.");
            }
        }
    }
}
