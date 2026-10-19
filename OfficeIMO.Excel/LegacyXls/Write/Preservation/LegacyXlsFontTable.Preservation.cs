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
                table._reservedPaletteSlots.Add(index);
                string color = table._paletteColors[index];
                if (!table._customColorIndexes.ContainsKey(color)) table._customColorIndexes.Add(color, checked((ushort)(index + 8)));
            }
            return table;
        }

    }
}
