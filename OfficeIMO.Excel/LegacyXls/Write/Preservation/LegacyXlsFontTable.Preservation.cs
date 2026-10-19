using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel.LegacyXls.Biff;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Globalization;
using System.Text;

namespace OfficeIMO.Excel.LegacyXls.Write {
    internal sealed partial class LegacyXlsFontTable {
        // Existing records are retained verbatim; this table reserves their indices for newly encoded cells.
        internal static LegacyXlsFontTable CreatePreserving(ExcelDocument document, LegacyXlsWorkbook source) {
            var table = new LegacyXlsFontTable(LegacyXlsThemeColorResolver.Create(document), false);
            foreach (var font in source.Fonts) {
                var key = new FontKey(font.Name ?? "Calibri", ToTwips(font.Size ?? 11), font.ColorIndex,
                    font.Bold, font.Italic, font.UnderlineStyle, font.Strikeout, font.Outline, font.Shadow,
                    font.Condense, font.Extend, font.Escapement, font.Family, font.CharacterSet);
                table._entries.Add(key);
                if (!table._indexes.ContainsKey(key)) table._indexes.Add(key, font.FontIndex);
                if (!table._richTextIndexes.ContainsKey(key)) table._richTextIndexes.Add(key, font.FontIndex);
            }
            table._paletteColors = BiffColorPalette.DefaultPaletteColors.ToList();
            for (int index = 0; index < source.PaletteColors.Count; index++) table._paletteColors[index] = source.PaletteColors[index];
            for (int index = 0; index < table._paletteColors.Count; index++) {
                string color = table._paletteColors[index];
                if (!table._customColorIndexes.ContainsKey(color)) table._customColorIndexes.Add(color, checked((ushort)(index + 8)));
            }
            void Reserve(ushort? index) { if (index is >= 8 and < 64) table._reservedPaletteSlots.Add(index.Value - 8); }
            foreach (var font in source.Fonts) Reserve(font.ColorIndex);
            foreach (var format in source.CellFormats) {
                Reserve(format.FillForegroundColorIndex); Reserve(format.FillBackgroundColorIndex);
                if (format.Border is not { } border) continue;
                Reserve(border.LeftColorIndex); Reserve(border.RightColorIndex); Reserve(border.TopColorIndex);
                Reserve(border.BottomColorIndex); Reserve(border.DiagonalColorIndex);
            }
            foreach (var sheet in source.Worksheets) { Reserve(sheet.GridLineColorIndex); Reserve(sheet.TabColorIndex); }
            // Opaque native features may contain palette references outside the typed style model.
            // Keep those slots rather than recoloring retained artwork or conditional formatting.
            if (source.DrawingRecords.Count != 0 || source.ChartRecords.Count != 0
                || source.PreservedFeatureRecords.Any(feature => feature.Kind is LegacyXlsUnsupportedFeatureKind.UnsupportedRecord
                    or LegacyXlsUnsupportedFeatureKind.StyleExtension or LegacyXlsUnsupportedFeatureKind.TableStyle
                    or LegacyXlsUnsupportedFeatureKind.ConditionalFormatting or LegacyXlsUnsupportedFeatureKind.FeatureExtension)
                || source.Worksheets.Any(sheet => sheet.ConditionalFormattings.Count != 0 || sheet.ConditionalFormattingExtensions.Count != 0))
                for (int index = 0; index < table._paletteColors.Count; index++) table._reservedPaletteSlots.Add(index);
            return table;
        }

    }
}
