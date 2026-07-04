using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal static class BiffStyleReader {
        internal static bool TryRead(
            BiffRecord record,
            LegacyXlsWorkbook workbook,
            List<LegacyXlsImportDiagnostic> diagnostics,
            out LegacyXlsCellStyleExtension? styleExtension) {
            styleExtension = null;
            switch ((BiffRecordType)record.Type) {
                case BiffRecordType.Style:
                    if (TryReadStyle(record, diagnostics, out LegacyXlsCellStyle? style)) {
                        workbook.AddCellStyle(style!);
                    }

                    return true;

                case BiffRecordType.XfCrc:
                    if (TryReadXfCrc(record, diagnostics, out LegacyXlsCellStyleExtension? xfCrcExtension)) {
                        workbook.AddCellStyleExtension(xfCrcExtension!);
                        styleExtension = xfCrcExtension;
                    }

                    return true;

                case BiffRecordType.XfExt:
                    if (TryReadXfExt(record, diagnostics, out LegacyXlsCellStyleExtension? extension)) {
                        workbook.AddCellStyleExtension(extension!);
                        styleExtension = extension;
                    }

                    return true;

                case BiffRecordType.StyleExt:
                    if (TryReadStyleExt(record, workbook, diagnostics, out LegacyXlsCellStyleExtension? styleExtExtension)) {
                        workbook.AddCellStyleExtension(styleExtExtension!);
                        styleExtension = styleExtExtension;
                    }

                    return true;

                default:
                    return false;
            }
        }

        private static bool TryReadStyle(BiffRecord record, List<LegacyXlsImportDiagnostic> diagnostics, out LegacyXlsCellStyle? style) {
            if (record.Payload.Length < 2) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-STYLE-SHORT",
                    "The Style record is shorter than expected.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                style = null;
                return false;
            }

            ushort flags = BiffRecordReader.ReadUInt16(record.Payload, 0);
            ushort styleFormatIndex = (ushort)(flags & 0x0fff);
            bool isBuiltIn = (flags & 0x8000) != 0;
            if (isBuiltIn) {
                if (record.Payload.Length < 4) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-STYLE-BUILTIN-SHORT",
                        "The built-in Style record is missing BuiltInStyle data.",
                        recordOffset: record.Offset,
                        recordType: record.Type));
                    style = null;
                    return false;
                }

                style = new LegacyXlsCellStyle(
                    styleFormatIndex,
                    isBuiltIn: true,
                    builtInStyleId: record.Payload[2],
                    outlineLevel: record.Payload[3],
                    name: null,
                    record.Offset,
                    record.Type);
                return true;
            }

            if (record.Payload.Length == 2) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-STYLE-NAME-MISSING",
                    "The custom Style record is missing its style name.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                style = null;
                return false;
            }

            try {
                int offset = 2;
                string name = ReadUnicodeOrShortByteString(record.Payload, ref offset);
                style = new LegacyXlsCellStyle(
                    styleFormatIndex,
                    isBuiltIn: false,
                    builtInStyleId: null,
                    outlineLevel: null,
                    name,
                    record.Offset,
                    record.Type);
                return true;
            } catch (InvalidDataException ex) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-STYLE-NAME-INVALID",
                    $"The custom Style record name could not be read. {ex.Message}",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                style = null;
                return false;
            }
        }

        private static string ReadUnicodeOrShortByteString(byte[] payload, ref int offset) {
            int originalOffset = offset;
            try {
                return BiffStringReader.ReadUnicodeString(payload, ref offset);
            } catch (InvalidDataException) {
                offset = originalOffset;
                return BiffStringReader.ReadShortByteString(payload, ref offset);
            }
        }

        private static bool TryReadXfCrc(BiffRecord record, List<LegacyXlsImportDiagnostic> diagnostics, out LegacyXlsCellStyleExtension? extension) {
            if (record.Payload.Length < 20) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-XFCRC-SHORT",
                    "The XFCRC record is shorter than expected.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                extension = null;
                return false;
            }

            ushort headerRecordType = BiffRecordReader.ReadUInt16(record.Payload, 0);
            if (headerRecordType != (ushort)BiffRecordType.XfCrc) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-XFCRC-HEADER-UNEXPECTED",
                    $"The XFCRC future record header declares record type 0x{headerRecordType:X4}.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
            }

            ushort reserved = BiffRecordReader.ReadUInt16(record.Payload, 12);
            if (reserved != 0) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-XFCRC-RESERVED-VALUE-UNEXPECTED",
                    "The XFCRC record contains a non-zero reserved field.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
            }

            extension = new LegacyXlsCellStyleExtension(
                "XFCRC",
                BiffRecordReader.ReadUInt16(record.Payload, 14),
                BiffRecordReader.ReadUInt32(record.Payload, 16),
                record.Offset,
                record.Type,
                record.Payload.Length);
            return true;
        }

        private static bool TryReadXfExt(BiffRecord record, List<LegacyXlsImportDiagnostic> diagnostics, out LegacyXlsCellStyleExtension? extension) {
            if (record.Payload.Length < 20) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-XFEXT-SHORT",
                    "The XFExt record is shorter than expected.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                extension = null;
                return false;
            }

            ushort headerRecordType = BiffRecordReader.ReadUInt16(record.Payload, 0);
            if (headerRecordType != (ushort)BiffRecordType.XfExt) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-XFEXT-HEADER-UNEXPECTED",
                    $"The XFExt future record header declares record type 0x{headerRecordType:X4}.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
            }

            ushort reserved1 = BiffRecordReader.ReadUInt16(record.Payload, 12);
            ushort reserved2 = BiffRecordReader.ReadUInt16(record.Payload, 16);
            if (reserved1 != 0 || reserved2 != 0) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-XFEXT-RESERVED-VALUE-UNEXPECTED",
                    "The XFExt record contains non-zero reserved fields.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
            }

            extension = new LegacyXlsCellStyleExtension(
                BiffRecordReader.ReadUInt16(record.Payload, 14),
                BiffRecordReader.ReadUInt16(record.Payload, 18),
                record.Offset,
                record.Type,
                record.Payload.Length,
                ReadXfExtProperties(record, diagnostics, BiffRecordReader.ReadUInt16(record.Payload, 18), 20, usesStyleXfPropMapping: false, out _));
            return true;
        }

        private static IReadOnlyList<LegacyXlsCellStyleExtensionProperty> ReadXfExtProperties(
            BiffRecord record,
            List<LegacyXlsImportDiagnostic> diagnostics,
            ushort extensionCount,
            int startOffset,
            bool usesStyleXfPropMapping,
            out bool fullyRead) {
            fullyRead = true;
            if (extensionCount == 0) {
                return Array.Empty<LegacyXlsCellStyleExtensionProperty>();
            }

            var properties = new List<LegacyXlsCellStyleExtensionProperty>(extensionCount);
            int offset = startOffset;
            for (int index = 0; index < extensionCount; index++) {
                if (offset + 4 > record.Payload.Length) {
                    fullyRead = false;
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-XFEXT-PROPERTY-SHORT",
                        "The XFExt record ended before all declared property extension headers could be read.",
                        recordOffset: record.Offset,
                        recordType: record.Type));
                    break;
                }

                ushort propertyType = BiffRecordReader.ReadUInt16(record.Payload, offset);
                ushort totalByteCount = BiffRecordReader.ReadUInt16(record.Payload, offset + 2);
                if (totalByteCount < 4) {
                    fullyRead = false;
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-XFEXT-PROPERTY-SIZE-INVALID",
                        $"An XFExt property extension declares an invalid size of {totalByteCount} bytes.",
                        recordOffset: record.Offset,
                        recordType: record.Type));
                    break;
                }

                if (offset + totalByteCount > record.Payload.Length) {
                    fullyRead = false;
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-XFEXT-PROPERTY-TRUNCATED",
                        "An XFExt property extension extends beyond the end of the record.",
                        recordOffset: record.Offset,
                        recordType: record.Type));
                    break;
                }

                int dataOffset = offset + 4;
                int dataByteCount = totalByteCount - 4;
                properties.Add(ReadXfExtProperty(
                    record.Payload,
                    dataOffset,
                    index,
                    propertyType,
                    totalByteCount,
                    dataByteCount,
                    usesStyleXfPropMapping));
                offset += totalByteCount;
            }

            if (offset < record.Payload.Length) {
                fullyRead = false;
            }

            return properties;
        }

        private static IReadOnlyList<LegacyXlsCellStyleExtensionProperty> ReadStyleExtProperties(
            BiffRecord record,
            List<LegacyXlsImportDiagnostic> diagnostics,
            int xfPropsOffset,
            out bool fullyRead) {
            fullyRead = true;
            if (xfPropsOffset >= record.Payload.Length) {
                return Array.Empty<LegacyXlsCellStyleExtensionProperty>();
            }

            if (xfPropsOffset + 4 > record.Payload.Length) {
                fullyRead = false;
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-STYLEEXT-XFPROPS-SHORT",
                    "The StyleExt record ended before its XFProps header could be read.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return Array.Empty<LegacyXlsCellStyleExtensionProperty>();
            }

            ushort propertyCount = BiffRecordReader.ReadUInt16(record.Payload, xfPropsOffset + 2);
            if (propertyCount > 1024) {
                fullyRead = false;
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-STYLEEXT-XFPROPS-COUNT-INVALID",
                    $"The StyleExt XFProps structure declares {propertyCount} properties.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                return Array.Empty<LegacyXlsCellStyleExtensionProperty>();
            }

            return ReadXfExtProperties(record, diagnostics, propertyCount, xfPropsOffset + 4, usesStyleXfPropMapping: true, out fullyRead);
        }

        private static LegacyXlsCellStyleExtensionProperty ReadXfExtProperty(
            byte[] payload,
            int dataOffset,
            int index,
            ushort propertyType,
            ushort totalByteCount,
            int dataByteCount,
            bool usesStyleXfPropMapping) {
            ushort? numericValue = null;
            string? numericValueName = null;
            ushort? colorType = null;
            string? colorTypeName = null;
            short? colorTintShade = null;
            uint? colorValue = null;
            ushort? borderStyle = null;
            string? borderStyleName = null;

            if (usesStyleXfPropMapping && IsStyleXfPropColorProperty(propertyType) && dataByteCount >= 8) {
                ReadStyleXfPropColor(payload, dataOffset, out colorType, out colorTypeName, out colorTintShade, out colorValue);
            } else if (usesStyleXfPropMapping && IsStyleXfPropBorderProperty(propertyType) && dataByteCount >= 10) {
                ReadStyleXfPropColor(payload, dataOffset, out colorType, out colorTypeName, out colorTintShade, out colorValue);
                borderStyle = BiffRecordReader.ReadUInt16(payload, dataOffset + 8);
                borderStyleName = GetBorderStyleName(borderStyle.Value);
            } else if (usesStyleXfPropMapping && propertyType == 0x0000 && dataByteCount >= 1) {
                numericValue = payload[dataOffset];
                numericValueName = GetFillPatternName(numericValue.Value);
            } else if (usesStyleXfPropMapping && propertyType == 0x0012 && dataByteCount >= 2) {
                numericValue = BiffRecordReader.ReadUInt16(payload, dataOffset);
                numericValueName = $"Indent:{numericValue.Value}";
            } else if (usesStyleXfPropMapping && propertyType == 0x0025 && dataByteCount >= 1) {
                numericValue = dataByteCount >= 2
                    ? BiffRecordReader.ReadUInt16(payload, dataOffset)
                    : payload[dataOffset];
                numericValueName = GetFontSchemeName(numericValue.Value);
            } else if (!usesStyleXfPropMapping && IsFullColorExtProperty(propertyType) && dataByteCount >= 8) {
                colorType = BiffRecordReader.ReadUInt16(payload, dataOffset);
                colorTypeName = GetXfExtColorTypeName(colorType.Value);
                colorTintShade = BiffRecordReader.ReadInt16(payload, dataOffset + 2);
                colorValue = BiffRecordReader.ReadUInt32(payload, dataOffset + 4);
            } else if (!usesStyleXfPropMapping && propertyType == 0x000E && dataByteCount >= 1) {
                numericValue = dataByteCount >= 2
                    ? BiffRecordReader.ReadUInt16(payload, dataOffset)
                    : payload[dataOffset];
                numericValueName = GetFontSchemeName(numericValue.Value);
            } else if (!usesStyleXfPropMapping && propertyType == 0x000F && dataByteCount >= 2) {
                numericValue = BiffRecordReader.ReadUInt16(payload, dataOffset);
                numericValueName = $"Indent:{numericValue.Value}";
            }

            return new LegacyXlsCellStyleExtensionProperty(
                index,
                propertyType,
                usesStyleXfPropMapping ? GetStyleXfPropTypeName(propertyType) : GetXfExtPropertyTypeName(propertyType),
                totalByteCount,
                dataByteCount,
                numericValue,
                numericValueName,
                colorType,
                colorTypeName,
                colorTintShade,
                colorValue,
                usesStyleXfPropMapping,
                borderStyle,
                borderStyleName);
        }

        private static bool TryReadStyleExt(
            BiffRecord record,
            LegacyXlsWorkbook workbook,
            List<LegacyXlsImportDiagnostic> diagnostics,
            out LegacyXlsCellStyleExtension? extension) {
            if (record.Payload.Length < 16) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-STYLEEXT-SHORT",
                    "The StyleExt record is shorter than expected.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
                extension = null;
                return false;
            }

            ushort headerRecordType = BiffRecordReader.ReadUInt16(record.Payload, 0);
            if (headerRecordType != (ushort)BiffRecordType.StyleExt) {
                diagnostics.Add(new LegacyXlsImportDiagnostic(
                    LegacyXlsDiagnosticSeverity.Warning,
                    "XLS-BIFF-STYLEEXT-HEADER-UNEXPECTED",
                    $"The StyleExt future record header declares record type 0x{headerRecordType:X4}.",
                    recordOffset: record.Offset,
                    recordType: record.Type));
            }

            byte flags = record.Payload[12];
            bool isBuiltIn = (flags & 0x01) != 0;
            bool isHidden = (flags & 0x02) != 0;
            bool isCustom = (flags & 0x04) != 0;
            byte category = record.Payload[13];
            ushort builtInData = BiffRecordReader.ReadUInt16(record.Payload, 14);
            string? styleName = null;
            int offset = 16;
            if (offset < record.Payload.Length) {
                try {
                    styleName = BiffStringReader.ReadWideString(record.Payload, ref offset);
                } catch (InvalidDataException ex) {
                    diagnostics.Add(new LegacyXlsImportDiagnostic(
                        LegacyXlsDiagnosticSeverity.Warning,
                        "XLS-BIFF-STYLEEXT-NAME-INVALID",
                        $"The StyleExt record name could not be read. {ex.Message}",
                        recordOffset: record.Offset,
                        recordType: record.Type));
                }
            }

            IReadOnlyList<LegacyXlsCellStyleExtensionProperty> properties = ReadStyleExtProperties(
                record,
                diagnostics,
                offset,
                out bool stylePropertiesFullyRead);
            bool hasUnparsedStyleProperties = !stylePropertiesFullyRead;
            ushort? associatedStyleFormatIndex = workbook.CellStyles.LastOrDefault()?.StyleFormatIndex;
            extension = new LegacyXlsCellStyleExtension(
                "StyleExt",
                isBuiltIn,
                isHidden,
                isCustom,
                category,
                GetStyleExtCategoryName(category),
                builtInData,
                styleName,
                associatedStyleFormatIndex,
                hasUnparsedStyleProperties,
                properties,
                record.Offset,
                record.Type,
                record.Payload.Length);
            return true;
        }

        private static string GetStyleExtCategoryName(byte category) {
            return category switch {
                0x00 => "Custom",
                0x01 => "GoodBadNeutral",
                0x02 => "DataModel",
                0x03 => "TitleAndHeading",
                0x04 => "ThemedCell",
                0x05 => "NumberFormat",
                _ => $"Unknown:0x{category:X2}"
            };
        }

        private static string GetXfExtPropertyTypeName(ushort propertyType) {
            return propertyType switch {
                0x0004 => "FillForegroundColor",
                0x0005 => "FillBackgroundColor",
                0x0006 => "FillGradient",
                0x0007 => "TopBorderColor",
                0x0008 => "BottomBorderColor",
                0x0009 => "LeftBorderColor",
                0x000A => "RightBorderColor",
                0x000B => "DiagonalBorderColor",
                0x000D => "TextColor",
                0x000E => "FontScheme",
                0x000F => "Indentation",
                _ => $"Unknown:0x{propertyType:X4}"
            };
        }

        private static string GetStyleXfPropTypeName(ushort propertyType) {
            return propertyType switch {
                0x0000 => "FillPattern",
                0x0001 => "FillForegroundColor",
                0x0002 => "FillBackgroundColor",
                0x0003 => "FillGradient",
                0x0004 => "FillGradientStop",
                0x0005 => "TextColor",
                0x0006 => "TopBorder",
                0x0007 => "BottomBorder",
                0x0008 => "LeftBorder",
                0x0009 => "RightBorder",
                0x000A => "DiagonalBorder",
                0x000B => "VerticalBorder",
                0x000C => "HorizontalBorder",
                0x0012 => "Indentation",
                0x0025 => "FontScheme",
                _ => $"Unknown:0x{propertyType:X4}"
            };
        }

        private static bool IsFullColorExtProperty(ushort propertyType) {
            return propertyType == 0x0004
                || propertyType == 0x0005
                || propertyType == 0x0007
                || propertyType == 0x0008
                || propertyType == 0x0009
                || propertyType == 0x000A
                || propertyType == 0x000B
                || propertyType == 0x000D;
        }

        private static bool IsStyleXfPropColorProperty(ushort propertyType) {
            return propertyType == 0x0001
                || propertyType == 0x0002
                || propertyType == 0x0005;
        }

        private static bool IsStyleXfPropBorderProperty(ushort propertyType) {
            return propertyType >= 0x0006 && propertyType <= 0x000A;
        }

        private static void ReadStyleXfPropColor(
            byte[] payload,
            int dataOffset,
            out ushort? colorType,
            out string? colorTypeName,
            out short? colorTintShade,
            out uint? colorValue) {
            byte flagsAndType = payload[dataOffset];
            colorType = checked((ushort)(flagsAndType >> 1));
            colorTypeName = GetXfExtColorTypeName(colorType.Value);
            byte indexedOrThemeValue = payload[dataOffset + 1];
            colorTintShade = BiffRecordReader.ReadInt16(payload, dataOffset + 2);
            colorValue = colorType.Value == 0x0001 || colorType.Value == 0x0003
                ? indexedOrThemeValue
                : BiffRecordReader.ReadUInt32(payload, dataOffset + 4);
        }

        private static string GetXfExtColorTypeName(ushort colorType) {
            return colorType switch {
                0x0000 => "Automatic",
                0x0001 => "Indexed",
                0x0002 => "Rgb",
                0x0003 => "Theme",
                0x0004 => "Ninch",
                _ => $"Unknown:0x{colorType:X4}"
            };
        }

        private static string GetFillPatternName(ushort pattern) {
            return pattern switch {
                0x0000 => "None",
                0x0001 => "Solid",
                0x0002 => "MediumGray",
                0x0003 => "DarkGray",
                0x0004 => "LightGray",
                0x0005 => "DarkHorizontal",
                0x0006 => "DarkVertical",
                0x0007 => "DarkDown",
                0x0008 => "DarkUp",
                0x0009 => "DarkGrid",
                0x000A => "DarkTrellis",
                0x000B => "LightHorizontal",
                0x000C => "LightVertical",
                0x000D => "LightDown",
                0x000E => "LightUp",
                0x000F => "LightGrid",
                0x0010 => "LightTrellis",
                0x0011 => "Gray125",
                0x0012 => "Gray0625",
                _ => $"Unknown:0x{pattern:X4}"
            };
        }

        private static string GetBorderStyleName(ushort style) {
            return style switch {
                0x0000 => "None",
                0x0001 => "Thin",
                0x0002 => "Medium",
                0x0003 => "Dashed",
                0x0004 => "Dotted",
                0x0005 => "Thick",
                0x0006 => "Double",
                0x0007 => "Hair",
                0x0008 => "MediumDashed",
                0x0009 => "DashDot",
                0x000A => "MediumDashDot",
                0x000B => "DashDotDot",
                0x000C => "MediumDashDotDot",
                0x000D => "SlantDashDot",
                _ => $"Unknown:0x{style:X4}"
            };
        }

        private static string GetFontSchemeName(ushort fontScheme) {
            return fontScheme switch {
                0x0000 => "None",
                0x0001 => "Major",
                0x0002 => "Minor",
                0x00ff => "Ninch",
                _ => $"Unknown:0x{fontScheme:X4}"
            };
        }
    }
}
