using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls.Projection {
    internal static partial class LegacyXlsWorkbookProjector {
        private static void ProjectCellStyles(LegacyXlsWorkbook workbook, ExcelDocument document) {
            LegacyXlsCellStyleExtension[] styleExtensions = workbook.CellStyleExtensions
                .Where(extension => extension.HasProjectableStyleMetadata)
                .ToArray();
            if (workbook.CellStyles.Count == 0 && styleExtensions.Length == 0) {
                return;
            }

            document.EnsureWorkbookThemeAndStyles();
            Stylesheet stylesheet = document.WorkbookPartRoot.WorkbookStylesPart!.Stylesheet!;
            CellStyles cellStyles = stylesheet.CellStyles ??= new CellStyles();

            var projectedStyleNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (LegacyXlsCellStyleExtension extension in styleExtensions) {
                ProjectStyleExtension(workbook, stylesheet, cellStyles, extension, projectedStyleNames);
            }

            foreach (LegacyXlsCellStyle style in workbook.CellStyles) {
                ProjectLegacyStyleRecord(workbook, stylesheet, cellStyles, style, projectedStyleNames);
            }

            stylesheet.CellStyleFormats!.Count = (uint)stylesheet.CellStyleFormats.Count();
            cellStyles.Count = (uint)cellStyles.Count();
            stylesheet.Save();
        }

        private static void ProjectStyleExtension(
            LegacyXlsWorkbook workbook,
            Stylesheet stylesheet,
            CellStyles cellStyles,
            LegacyXlsCellStyleExtension extension,
            HashSet<string> projectedStyleNames) {
            string? styleName = NormalizeStyleName(extension.StyleName);
            if (styleName == null || !projectedStyleNames.Add(styleName)) {
                return;
            }

            uint formatId = ResolveCellStyleFormatId(workbook, stylesheet, extension.AssociatedStyleFormatIndex);
            CellStyle style = GetOrCreateCellStyle(cellStyles, styleName);
            style.Name = styleName;
            style.FormatId = formatId;
            ApplyStyleExtensionMetadata(style, extension);
        }

        private static void ProjectLegacyStyleRecord(
            LegacyXlsWorkbook workbook,
            Stylesheet stylesheet,
            CellStyles cellStyles,
            LegacyXlsCellStyle style,
            HashSet<string> projectedStyleNames) {
            string? styleName = NormalizeStyleName(style.Name);
            if (styleName == null || !projectedStyleNames.Add(styleName)) {
                return;
            }

            uint formatId = ResolveCellStyleFormatId(workbook, stylesheet, style.StyleFormatIndex);
            CellStyle cellStyle = GetOrCreateCellStyle(cellStyles, styleName);
            cellStyle.Name = styleName;
            cellStyle.FormatId = formatId;
            if (style.IsBuiltIn && style.BuiltInStyleId.HasValue) {
                cellStyle.BuiltinId = style.BuiltInStyleId.Value;
                if (style.OutlineLevel.HasValue) {
                    cellStyle.OutlineLevel = style.OutlineLevel.Value;
                }
            }
        }

        private static void ApplyStyleExtensionMetadata(CellStyle style, LegacyXlsCellStyleExtension extension) {
            if (extension.IsBuiltInStyle == true && extension.BuiltInData.HasValue) {
                style.BuiltinId = GetBuiltInStyleId(extension.BuiltInData.Value);
                byte outlineLevel = GetBuiltInStyleOutlineLevel(extension.BuiltInData.Value);
                if (outlineLevel != 0xff) {
                    style.OutlineLevel = outlineLevel;
                }
            }

            if (extension.IsHidden == true) {
                style.Hidden = true;
            }

            if (extension.IsCustom == true) {
                style.CustomBuiltin = true;
            }
        }

        private static uint ResolveCellStyleFormatId(LegacyXlsWorkbook workbook, Stylesheet stylesheet, ushort? legacyStyleFormatIndex) {
            if (!legacyStyleFormatIndex.HasValue) {
                return 0U;
            }

            LegacyXlsCellFormat? format = workbook.GetCellFormat(legacyStyleFormatIndex.Value);
            if (format == null) {
                return 0U;
            }

            CellFormat candidate = ExcelSheet.CreateLegacyCellFormat(stylesheet, workbook, format);
            return AppendOrReuseCellStyleFormat(stylesheet, candidate);
        }

        private static uint AppendOrReuseCellStyleFormat(Stylesheet stylesheet, CellFormat candidate) {
            CellStyleFormats formats = stylesheet.CellStyleFormats ??= new CellStyleFormats();
            CellFormat[] existingFormats = formats.Elements<CellFormat>().ToArray();
            string candidateXml = candidate.OuterXml;
            for (int i = 0; i < existingFormats.Length; i++) {
                if (string.Equals(existingFormats[i].OuterXml, candidateXml, StringComparison.Ordinal)) {
                    return checked((uint)i);
                }
            }

            formats.Append((CellFormat)candidate.CloneNode(true));
            formats.Count = (uint)formats.Count();
            return checked((uint)(formats.Count!.Value - 1));
        }

        private static CellStyle GetOrCreateCellStyle(CellStyles cellStyles, string styleName) {
            CellStyle? existing = cellStyles.Elements<CellStyle>()
                .FirstOrDefault(style => string.Equals(style.Name?.Value, styleName, StringComparison.OrdinalIgnoreCase));
            if (existing != null) {
                return existing;
            }

            var created = new CellStyle();
            cellStyles.Append(created);
            return created;
        }

        private static string? NormalizeStyleName(string? styleName) {
            return string.IsNullOrWhiteSpace(styleName) ? null : styleName!.Trim();
        }

        private static byte GetBuiltInStyleId(ushort builtInData) {
            return checked((byte)(builtInData & 0xff));
        }

        private static byte GetBuiltInStyleOutlineLevel(ushort builtInData) {
            return checked((byte)((builtInData >> 8) & 0xff));
        }
    }
}
