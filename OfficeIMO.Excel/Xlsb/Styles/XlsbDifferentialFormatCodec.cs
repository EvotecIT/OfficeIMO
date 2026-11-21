using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.Xlsb.Biff12;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Projection;
using OfficeIMO.Excel.Xlsb.Write;

namespace OfficeIMO.Excel.Xlsb.Styles {
    /// <summary>
    /// Encodes and decodes differential formatting (MS-XLSB 2.2.6.2) as a BrtDXF record (2.4.359) whose XFProps
    /// (2.5.164) carry one XFProp (2.5.159) per font, number format, fill, alignment, protection and border property.
    /// </summary>
    internal static class XlsbDifferentialFormatCodec {
        internal const int BrtDXF = 507;
        private const int MaximumFontNameLength = 32;

        // xfPropType values (MS-XLSB 2.5.159).
        private const ushort FillPattern = 0x00;
        private const ushort ForegroundColor = 0x01;
        private const ushort BackgroundColor = 0x02;
        private const ushort GradientFill = 0x03;
        private const ushort GradientStop = 0x04;
        private const ushort TextColor = 0x05;
        private const ushort TopBorder = 0x06;
        private const ushort BottomBorder = 0x07;
        private const ushort LeftBorder = 0x08;
        private const ushort RightBorder = 0x09;
        private const ushort DiagonalBorder = 0x0A;
        private const ushort VerticalBorder = 0x0B;
        private const ushort HorizontalBorder = 0x0C;
        private const ushort DiagonalUp = 0x0D;
        private const ushort DiagonalDown = 0x0E;
        private const ushort HorizontalAlignment = 0x0F;
        private const ushort VerticalAlignment = 0x10;
        private const ushort TextRotation = 0x11;
        private const ushort AbsoluteIndent = 0x12;
        private const ushort ReadingOrder = 0x13;
        private const ushort WrapText = 0x14;
        private const ushort JustifyDistributed = 0x15;
        private const ushort ShrinkToFit = 0x16;
        private const ushort FontName = 0x18;
        private const ushort FontBold = 0x19;
        private const ushort FontUnderline = 0x1A;
        private const ushort FontScript = 0x1B;
        private const ushort FontItalic = 0x1C;
        private const ushort FontStrike = 0x1D;
        private const ushort FontOutline = 0x1E;
        private const ushort FontShadow = 0x1F;
        private const ushort FontCondense = 0x20;
        private const ushort FontExtend = 0x21;
        private const ushort FontCharSet = 0x22;
        private const ushort FontFamily = 0x23;
        private const ushort FontSize = 0x24;
        private const ushort FontScheme = 0x25;
        private const ushort NumberFormatCode = 0x26;
        private const ushort NumberFormatId = 0x29;
        private const ushort RelativeIndent = 0x2A;
        private const ushort Locked = 0x2B;
        private const ushort Hidden = 0x2C;

        /// <summary>Encodes <paramref name="format"/> as a BrtDXF record or throws when it holds content XFProps cannot carry.</summary>
        internal static XlsbGeneratedRecord CreateRecord(DifferentialFormat format) {
            ArgumentNullException.ThrowIfNull(format);
            Require(format.ChildElements.All(child => child is Font or NumberingFormat or Fill or Alignment or Protection or Border)
                && XlsbOpenXmlAttributeValidator.FindUnsupported(format, Array.Empty<string>()) == null,
                "extension or unknown content");
            var properties = new List<(ushort Type, byte[] Data)>();
            foreach (OpenXmlElement child in format.ChildElements) {
                switch (child) {
                    case Font font: AddFont(properties, font); break;
                    case NumberingFormat numberFormat: AddNumberFormat(properties, numberFormat); break;
                    case Fill fill: AddFill(properties, fill); break;
                    case Alignment alignment: AddAlignment(properties, alignment); break;
                    case Protection protection: AddProtection(properties, protection); break;
                    case Border border: AddBorder(properties, border); break;
                }
            }
            if (properties.Count > ushort.MaxValue) Require(false, "more properties than XFProps can count");
            // Excel writes the properties by type, with the number format code after its id.
            properties = properties
                .Select((property, index) => (property, index))
                .OrderBy(item => item.property.Type == NumberFormatCode ? NumberFormatId + 0.5 : item.property.Type)
                .ThenBy(item => item.index)
                .Select(item => item.property)
                .ToList();

            using var payload = new MemoryStream();
            // The first two bytes hold 15 unused bits and fNewBorder in the most significant bit; fNewBorder allows the
            // internal (vertical and horizontal) borders and is always set, as Excel writes it.
            XlsbPayloadWriter.WriteUInt16(payload, 0x8000);
            XlsbPayloadWriter.WriteUInt16(payload, 0);
            XlsbPayloadWriter.WriteUInt16(payload, (ushort)properties.Count);
            foreach ((ushort type, byte[] data) in properties) {
                if (data.Length > ushort.MaxValue - 4) Require(false, "a property longer than an XFProp can hold");
                XlsbPayloadWriter.WriteUInt16(payload, type);
                XlsbPayloadWriter.WriteUInt16(payload, (ushort)(data.Length + 4));
                payload.Write(data, 0, data.Length);
            }
            return new XlsbGeneratedRecord(BrtDXF, payload.ToArray());
        }

        private static void AddFont(List<(ushort, byte[])> properties, Font font) {
            Require(XlsbOpenXmlAttributeValidator.FindUnsupported(font, Array.Empty<string>()) == null, "font attributes");
            foreach (OpenXmlElement child in font.ChildElements) {
                switch (child) {
                    case Bold bold:
                        properties.Add((FontBold, UInt16(ExcelOpenXmlFontProperty.IsEnabled(bold) ? (ushort)0x02BC : (ushort)0x0190)));
                        break;
                    case Italic value: properties.Add((FontItalic, Flag(value))); break;
                    case Strike value: properties.Add((FontStrike, Flag(value))); break;
                    case Outline value: properties.Add((FontOutline, Flag(value))); break;
                    case Shadow value: properties.Add((FontShadow, Flag(value))); break;
                    case Condense value: properties.Add((FontCondense, Flag(value))); break;
                    case Extend value: properties.Add((FontExtend, Flag(value))); break;
                    case Underline underline:
                        properties.Add((FontUnderline, UInt16(ToUnderline(underline.Val?.Value ?? UnderlineValues.Single))));
                        break;
                    case VerticalTextAlignment script:
                        properties.Add((FontScript, UInt16(ToScript(script.Val?.Value))));
                        break;
                    case DocumentFormat.OpenXml.Spreadsheet.FontSize size: {
                        double points = size.Val?.Value ?? 0D;
                        double twips = Math.Round(points * 20D, MidpointRounding.AwayFromZero);
                        Require(twips >= 20D && twips <= 8191D, $"font size {points}");
                        properties.Add((FontSize, UInt32((uint)twips)));
                        break;
                    }
                    case DocumentFormat.OpenXml.Spreadsheet.Color color:
                        properties.Add((TextColor, Color(color)));
                        break;
                    case DocumentFormat.OpenXml.Spreadsheet.FontName name: {
                        string text = name.Val?.Value ?? string.Empty;
                        Require(text.Length > 0 && text.Length <= MaximumFontNameLength, $"font name '{text}'");
                        using var data = new MemoryStream(2 + text.Length * 2);
                        XlsbPayloadWriter.WriteUInt16(data, (ushort)text.Length);
                        foreach (char character in text) XlsbPayloadWriter.WriteUInt16(data, character);
                        properties.Add((FontName, data.ToArray()));
                        break;
                    }
                    case FontFamilyNumbering family: {
                        int value = family.Val?.Value ?? 0;
                        Require(value >= 0 && value <= 5, $"font family {value}");
                        properties.Add((FontFamily, new[] { (byte)value }));
                        break;
                    }
                    case DocumentFormat.OpenXml.Spreadsheet.FontCharSet charset: {
                        int value = charset.Val?.Value ?? 0;
                        Require(value >= 0 && value <= byte.MaxValue, $"font character set {value}");
                        properties.Add((FontCharSet, new[] { (byte)value }));
                        break;
                    }
                    case DocumentFormat.OpenXml.Spreadsheet.FontScheme scheme:
                        properties.Add((FontScheme, new[] { ToFontScheme(scheme.Val?.Value) }));
                        break;
                    default:
                        Require(false, $"font element '{child.LocalName}'");
                        break;
                }
            }
        }

        private static void AddNumberFormat(List<(ushort, byte[])> properties, NumberingFormat format) {
            Require(XlsbOpenXmlAttributeValidator.FindUnsupported(format, new[] { "numFmtId", "formatCode" }) == null && !format.HasChildren,
                "number format content");
            uint id = format.NumberFormatId?.Value ?? 0U;
            Require(id <= ushort.MaxValue, $"number format id {id}");
            string? code = format.FormatCode?.Value;
            properties.Add((NumberFormatId, UInt16((ushort)id)));
            if (!string.IsNullOrEmpty(code)) {
                Require(code!.Length <= 255, "a number format longer than 255 characters");
                properties.Add((NumberFormatCode, NumberFormatString(code)));
            }
        }

        private static void AddFill(List<(ushort, byte[])> properties, Fill fill) {
            Require(XlsbOpenXmlAttributeValidator.FindUnsupported(fill, Array.Empty<string>()) == null
                && fill.ChildElements.Count == 1 && fill.ChildElements[0] is PatternFill, "a gradient or extended fill");
            var pattern = (PatternFill)fill.ChildElements[0];
            Require(XlsbOpenXmlAttributeValidator.FindUnsupported(pattern, new[] { "patternType" }) == null
                && pattern.ChildElements.All(child => child is DocumentFormat.OpenXml.Spreadsheet.ForegroundColor or DocumentFormat.OpenXml.Spreadsheet.BackgroundColor),
                "pattern fill content");
            if (pattern.PatternType?.Value is PatternValues patternType) {
                properties.Add((FillPattern, new[] { (byte)XlsbStylesheetPartWriter.ToFillPattern(patternType) }));
            }
            if (pattern.ForegroundColor != null) properties.Add((ForegroundColor, Color(pattern.ForegroundColor)));
            if (pattern.BackgroundColor != null) properties.Add((BackgroundColor, Color(pattern.BackgroundColor)));
        }

        private static void AddAlignment(List<(ushort, byte[])> properties, Alignment alignment) {
            Require(!alignment.HasChildren && XlsbOpenXmlAttributeValidator.FindUnsupported(alignment, new[] {
                "horizontal", "vertical", "textRotation", "wrapText", "indent", "relativeIndent", "justifyLastLine", "shrinkToFit", "readingOrder"
            }) == null, "alignment content");
            if (alignment.Horizontal?.Value is HorizontalAlignmentValues horizontal) {
                properties.Add((HorizontalAlignment, new[] { XlsbStylesheetPartWriter.ToHorizontalAlignment(horizontal) }));
            }
            if (alignment.Vertical?.Value is VerticalAlignmentValues vertical) {
                properties.Add((VerticalAlignment, new[] { XlsbStylesheetPartWriter.ToVerticalAlignment(vertical) }));
            }
            if (alignment.TextRotation?.Value is uint rotation) {
                Require(rotation <= 180U || rotation == 255U, $"text rotation {rotation}");
                properties.Add((TextRotation, new[] { (byte)rotation }));
            }
            if (alignment.Indent?.Value is uint indent) {
                Require(indent <= 15U, $"indent {indent}");
                properties.Add((AbsoluteIndent, UInt16((ushort)indent)));
            }
            if (alignment.RelativeIndent?.Value is int relative) {
                Require(relative >= -15 && relative <= 15, $"relative indent {relative}");
                properties.Add((RelativeIndent, UInt16(unchecked((ushort)(short)relative))));
            }
            if (alignment.ReadingOrder?.Value is uint order) {
                Require(order <= 2U, $"reading order {order}");
                properties.Add((ReadingOrder, new[] { (byte)order }));
            }
            if (alignment.WrapText != null) properties.Add((WrapText, Flag(alignment.WrapText.Value)));
            if (alignment.JustifyLastLine != null) {
                Require(alignment.JustifyLastLine.Value != true || alignment.Horizontal?.Value == HorizontalAlignmentValues.Distributed,
                    "justifyLastLine without distributed horizontal alignment");
                properties.Add((JustifyDistributed, Flag(alignment.JustifyLastLine.Value)));
            }
            if (alignment.ShrinkToFit != null) properties.Add((ShrinkToFit, Flag(alignment.ShrinkToFit.Value)));
        }

        private static void AddProtection(List<(ushort, byte[])> properties, Protection protection) {
            Require(!protection.HasChildren && XlsbOpenXmlAttributeValidator.FindUnsupported(protection, new[] { "locked", "hidden" }) == null,
                "protection content");
            if (protection.Locked != null) properties.Add((Locked, Flag(protection.Locked.Value)));
            if (protection.Hidden != null) properties.Add((Hidden, Flag(protection.Hidden.Value)));
        }

        private static void AddBorder(List<(ushort, byte[])> properties, Border border) {
            Require(XlsbOpenXmlAttributeValidator.FindUnsupported(border, new[] { "diagonalUp", "diagonalDown", "outline" }) == null
                && border.Outline?.Value != false, "border attributes");
            foreach (OpenXmlElement child in border.ChildElements) {
                ushort type = child switch {
                    DocumentFormat.OpenXml.Spreadsheet.LeftBorder => LeftBorder,
                    DocumentFormat.OpenXml.Spreadsheet.RightBorder => RightBorder,
                    DocumentFormat.OpenXml.Spreadsheet.TopBorder => TopBorder,
                    DocumentFormat.OpenXml.Spreadsheet.BottomBorder => BottomBorder,
                    DocumentFormat.OpenXml.Spreadsheet.DiagonalBorder => DiagonalBorder,
                    DocumentFormat.OpenXml.Spreadsheet.VerticalBorder => VerticalBorder,
                    DocumentFormat.OpenXml.Spreadsheet.HorizontalBorder => HorizontalBorder,
                    _ => ushort.MaxValue
                };
                Require(type != ushort.MaxValue, $"border element '{child.LocalName}'");
                var side = (BorderPropertiesType)child;
                Require(XlsbOpenXmlAttributeValidator.FindUnsupported(side, new[] { "style" }) == null
                    && side.ChildElements.All(item => item is DocumentFormat.OpenXml.Spreadsheet.Color), "border side content");
                // A side without a color is written with the automatic color, as Excel writes an empty side.
                using var data = new MemoryStream(10);
                byte[] color = Color(side.Color);
                data.Write(color, 0, color.Length);
                XlsbPayloadWriter.WriteUInt16(data, XlsbStylesheetPartWriter.ToBorderStyle(side.Style?.Value));
                properties.Add((type, data.ToArray()));
            }
            if (border.DiagonalUp != null) properties.Add((DiagonalUp, Flag(border.DiagonalUp.Value)));
            if (border.DiagonalDown != null) properties.Add((DiagonalDown, Flag(border.DiagonalDown.Value)));
        }

        /// <summary>
        /// Decodes a BrtDXF payload into the canonical SpreadsheetML dxf. Properties that SpreadsheetML cannot hold (such as
        /// gradient fills) are reported through <paramref name="unprojected"/> by xfPropType.
        /// </summary>
        internal static DifferentialFormat Read(byte[] payload, Action<ushort> unprojected) {
            var cursor = new XlsbBinaryCursor(payload);
            cursor.Skip(2);
            cursor.Skip(2);
            ushort count = cursor.ReadUInt16();
            var font = new Font();
            NumberingFormat? numberFormat = null;
            PatternFill? pattern = null;
            Alignment? alignment = null;
            Protection? protection = null;
            Border? border = null;
            for (int index = 0; index < count; index++) {
                ushort type = cursor.ReadUInt16();
                ushort size = cursor.ReadUInt16();
                if (size < 4) throw new InvalidDataException("A BrtDXF property has an invalid size.");
                var data = new XlsbBinaryCursor(cursor.ReadBytes(size - 4));
                switch (type) {
                    case FillPattern: (pattern ??= new PatternFill()).PatternType = XlsbStylesheetProjector.ToFillPattern(data.ReadByte()); break;
                    case ForegroundColor:
                        if (ReadColor<DocumentFormat.OpenXml.Spreadsheet.ForegroundColor>(data) is { } foreground) (pattern ??= new PatternFill()).ForegroundColor = foreground;
                        break;
                    case BackgroundColor:
                        if (ReadColor<DocumentFormat.OpenXml.Spreadsheet.BackgroundColor>(data) is { } background) (pattern ??= new PatternFill()).BackgroundColor = background;
                        break;
                    case TextColor:
                        if (ReadColor<DocumentFormat.OpenXml.Spreadsheet.Color>(data) is { } textColor) font.Color = textColor;
                        break;
                    case TopBorder: (border ??= new Border()).TopBorder = ReadBorderSide<DocumentFormat.OpenXml.Spreadsheet.TopBorder>(data); break;
                    case BottomBorder: (border ??= new Border()).BottomBorder = ReadBorderSide<DocumentFormat.OpenXml.Spreadsheet.BottomBorder>(data); break;
                    case LeftBorder: (border ??= new Border()).LeftBorder = ReadBorderSide<DocumentFormat.OpenXml.Spreadsheet.LeftBorder>(data); break;
                    case RightBorder: (border ??= new Border()).RightBorder = ReadBorderSide<DocumentFormat.OpenXml.Spreadsheet.RightBorder>(data); break;
                    case DiagonalBorder: (border ??= new Border()).DiagonalBorder = ReadBorderSide<DocumentFormat.OpenXml.Spreadsheet.DiagonalBorder>(data); break;
                    case VerticalBorder: (border ??= new Border()).VerticalBorder = ReadBorderSide<DocumentFormat.OpenXml.Spreadsheet.VerticalBorder>(data); break;
                    case HorizontalBorder: (border ??= new Border()).HorizontalBorder = ReadBorderSide<DocumentFormat.OpenXml.Spreadsheet.HorizontalBorder>(data); break;
                    case DiagonalUp: (border ??= new Border()).DiagonalUp = data.ReadByte() != 0; break;
                    case DiagonalDown: (border ??= new Border()).DiagonalDown = data.ReadByte() != 0; break;
                    case HorizontalAlignment: (alignment ??= new Alignment()).Horizontal = FromHorizontalAlignment(data.ReadByte()); break;
                    case VerticalAlignment: (alignment ??= new Alignment()).Vertical = FromVerticalAlignment(data.ReadByte()); break;
                    case TextRotation: (alignment ??= new Alignment()).TextRotation = data.ReadByte(); break;
                    case AbsoluteIndent: (alignment ??= new Alignment()).Indent = data.ReadUInt16(); break;
                    case RelativeIndent: {
                        short relative = data.ReadInt16();
                        if (relative != 255) (alignment ??= new Alignment()).RelativeIndent = relative;
                        break;
                    }
                    case ReadingOrder: (alignment ??= new Alignment()).ReadingOrder = data.ReadByte(); break;
                    case WrapText: (alignment ??= new Alignment()).WrapText = data.ReadByte() != 0; break;
                    case JustifyDistributed: (alignment ??= new Alignment()).JustifyLastLine = data.ReadByte() != 0; break;
                    case ShrinkToFit: (alignment ??= new Alignment()).ShrinkToFit = data.ReadByte() != 0; break;
                    case FontName: {
                        ushort length = data.ReadUInt16();
                        font.FontName = new DocumentFormat.OpenXml.Spreadsheet.FontName { Val = System.Text.Encoding.Unicode.GetString(data.ReadBytes(length * 2)) };
                        break;
                    }
                    case FontBold: font.Bold = data.ReadUInt16() >= 0x02BC ? new Bold() : new Bold { Val = false }; break;
                    case FontUnderline: font.Underline = new Underline { Val = FromUnderline(data.ReadUInt16()) }; break;
                    case FontScript: font.VerticalTextAlignment = new VerticalTextAlignment { Val = FromScript(data.ReadUInt16()) }; break;
                    case FontItalic: font.Italic = data.ReadByte() != 0 ? new Italic() : new Italic { Val = false }; break;
                    case FontStrike: font.Strike = data.ReadByte() != 0 ? new Strike() : new Strike { Val = false }; break;
                    case FontOutline: font.Outline = data.ReadByte() != 0 ? new Outline() : new Outline { Val = false }; break;
                    case FontShadow: font.Shadow = data.ReadByte() != 0 ? new Shadow() : new Shadow { Val = false }; break;
                    case FontCondense: font.Condense = data.ReadByte() != 0 ? new Condense() : new Condense { Val = false }; break;
                    case FontExtend: font.Extend = data.ReadByte() != 0 ? new Extend() : new Extend { Val = false }; break;
                    case FontCharSet: font.AddChild(new DocumentFormat.OpenXml.Spreadsheet.FontCharSet { Val = data.ReadByte() }, true); break;
                    case FontFamily: font.FontFamilyNumbering = new FontFamilyNumbering { Val = data.ReadByte() }; break;
                    case FontSize: font.FontSize = new DocumentFormat.OpenXml.Spreadsheet.FontSize { Val = data.ReadUInt32() / 20D }; break;
                    case FontScheme: font.AddChild(new DocumentFormat.OpenXml.Spreadsheet.FontScheme { Val = FromFontScheme(data.ReadByte()) }, true); break;
                    case NumberFormatId: (numberFormat ??= new NumberingFormat()).NumberFormatId = data.ReadUInt16(); break;
                    case NumberFormatCode: (numberFormat ??= new NumberingFormat()).FormatCode = ReadNumberFormatString(data); break;
                    case Locked: (protection ??= new Protection()).Locked = data.ReadByte() != 0; break;
                    case Hidden: (protection ??= new Protection()).Hidden = data.ReadByte() != 0; break;
                    default:
                        unprojected(type);
                        break;
                }
            }
            if (cursor.Remaining != 0) throw new InvalidDataException("A BrtDXF record has unexpected trailing data.");

            // CT_Dxf order: font, numFmt, fill, alignment, protection, border.
            var format = new DifferentialFormat();
            if (font.HasChildren) format.Append(font);
            if (numberFormat != null) {
                numberFormat.NumberFormatId ??= 0U;
                numberFormat.FormatCode ??= string.Empty;
                format.Append(numberFormat);
            }
            if (pattern != null) format.Append(new Fill(pattern));
            if (alignment != null) format.Append(alignment);
            if (protection != null) format.Append(protection);
            if (border != null) format.Append(border);
            return format;
        }

        private static T? ReadColor<T>(XlsbBinaryCursor data) where T : ColorType, new() {
            XlsbColor color = XlsbStylesheetReader.ReadColor(data);
            if (color.Type == XlsbColorNotSet) return null;
            var result = new T();
            XlsbStylesheetProjector.ApplyColor(result, color);
            return result;
        }

        private static T ReadBorderSide<T>(XlsbBinaryCursor data) where T : BorderPropertiesType, new() {
            // An automatic color and no style are how an empty side is stored, so they project to an empty element.
            var side = new T();
            XlsbColor color = XlsbStylesheetReader.ReadColor(data);
            if (XlsbStylesheetProjector.ToBorderStyle((byte)data.ReadUInt16()) is BorderStyleValues style) side.Style = style;
            if (color.Type is not 0 and not XlsbColorNotSet) {
                var element = new DocumentFormat.OpenXml.Spreadsheet.Color();
                XlsbStylesheetProjector.ApplyColor(element, color);
                side.Color = element;
            }
            return side;
        }

        // XFPropColor xclrType 4 (MS-XLSB 2.5.161): the color is not set.
        private const byte XlsbColorNotSet = 4;

        private static byte[] Color(ColorType? color) {
            using var data = new MemoryStream(8);
            if (color == null) {
                // fValidRGBA with the automatic color type (MS-XLSB 2.5.161).
                return new byte[] { 0x01, 0x40, 0, 0, 0, 0, 0, 0 };
            }
            XlsbStylesheetPartWriter.WriteColor(data, color, allowFullShade: false);
            return data.ToArray();
        }

        // Excel stores the format code as a 2-byte character count followed by UTF-16 characters.
        private static byte[] NumberFormatString(string code) {
            using var data = new MemoryStream(2 + code.Length * 2);
            XlsbPayloadWriter.WriteUInt16(data, (ushort)code.Length);
            foreach (char character in code) XlsbPayloadWriter.WriteUInt16(data, character);
            return data.ToArray();
        }

        private static string ReadNumberFormatString(XlsbBinaryCursor data) {
            ushort length = data.ReadUInt16();
            return System.Text.Encoding.Unicode.GetString(data.ReadBytes(length * 2));
        }

        private static byte[] Flag(BooleanPropertyType value) => new[] { ExcelOpenXmlFontProperty.IsEnabled(value) ? (byte)1 : (byte)0 };

        private static byte[] Flag(bool value) => new[] { value ? (byte)1 : (byte)0 };

        private static byte[] UInt16(ushort value) => new[] { (byte)value, (byte)(value >> 8) };

        private static byte[] UInt32(uint value) => new[] { (byte)value, (byte)(value >> 8), (byte)(value >> 16), (byte)(value >> 24) };

        private static ushort ToUnderline(UnderlineValues value) {
            if (value == UnderlineValues.Single) return 1;
            if (value == UnderlineValues.Double) return 2;
            if (value == UnderlineValues.SingleAccounting) return 0x21;
            if (value == UnderlineValues.DoubleAccounting) return 0x22;
            return 0;
        }

        private static ushort ToScript(VerticalAlignmentRunValues? value) {
            if (value == VerticalAlignmentRunValues.Superscript) return 1;
            if (value == VerticalAlignmentRunValues.Subscript) return 2;
            return 0;
        }

        private static byte ToFontScheme(FontSchemeValues? value) {
            if (value == FontSchemeValues.Major) return 1;
            if (value == FontSchemeValues.Minor) return 2;
            return 0;
        }

        private static HorizontalAlignmentValues FromHorizontalAlignment(byte value) => value switch {
            1 => HorizontalAlignmentValues.Left,
            2 => HorizontalAlignmentValues.Center,
            3 => HorizontalAlignmentValues.Right,
            4 => HorizontalAlignmentValues.Fill,
            5 => HorizontalAlignmentValues.Justify,
            6 => HorizontalAlignmentValues.CenterContinuous,
            7 => HorizontalAlignmentValues.Distributed,
            _ => HorizontalAlignmentValues.General
        };

        private static VerticalAlignmentValues FromVerticalAlignment(byte value) => value switch {
            0 => VerticalAlignmentValues.Top,
            1 => VerticalAlignmentValues.Center,
            3 => VerticalAlignmentValues.Justify,
            4 => VerticalAlignmentValues.Distributed,
            _ => VerticalAlignmentValues.Bottom
        };

        private static UnderlineValues FromUnderline(ushort value) => value switch {
            1 => UnderlineValues.Single,
            2 => UnderlineValues.Double,
            0x21 => UnderlineValues.SingleAccounting,
            0x22 => UnderlineValues.DoubleAccounting,
            _ => UnderlineValues.None
        };

        private static VerticalAlignmentRunValues FromScript(ushort value) => value switch {
            1 => VerticalAlignmentRunValues.Superscript,
            2 => VerticalAlignmentRunValues.Subscript,
            _ => VerticalAlignmentRunValues.Baseline
        };

        private static FontSchemeValues FromFontScheme(byte value) => value switch {
            1 => FontSchemeValues.Major,
            2 => FontSchemeValues.Minor,
            _ => FontSchemeValues.None
        };

        private static void Require(bool condition, string description) {
            if (!condition) {
                throw new NotSupportedException($"Native XLSB generation cannot encode a differential format with {description}.");
            }
        }
    }
}
