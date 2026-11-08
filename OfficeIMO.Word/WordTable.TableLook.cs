using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word {
    public partial class WordTable {
        private const int NativeDefaultTableLook = 0x04A0;
        private const int ExpandedDefaultTableLook = 0x07E0;
        private const int TableLookFirstRow = 0x0020;
        private const int TableLookLastRow = 0x0040;
        private const int TableLookFirstColumn = 0x0080;
        private const int TableLookLastColumn = 0x0100;
        private const int TableLookNoHorizontalBand = 0x0200;
        private const int TableLookNoVerticalBand = 0x0400;

        private static TableLook CreateTableLook(bool firstRow, bool lastRow, bool firstColumn, bool lastColumn, bool noHorizontalBand, bool noVerticalBand) {
            var tableLook = new TableLook();
            var mask = 0;
            if (firstRow) mask |= TableLookFirstRow;
            if (lastRow) mask |= TableLookLastRow;
            if (firstColumn) mask |= TableLookFirstColumn;
            if (lastColumn) mask |= TableLookLastColumn;
            if (noHorizontalBand) mask |= TableLookNoHorizontalBand;
            if (noVerticalBand) mask |= TableLookNoVerticalBand;
            SetTableLookMask(tableLook, mask);
            return tableLook;
        }

        /// <summary>Reads the native first-row role without constructing another table wrapper.</summary>
        internal static bool HasFirstRowFormatting(Table table) {
            TableLook? look = table.GetFirstChild<TableProperties>()?.TableLook;
            return (GetTableLookMask(look) & TableLookFirstRow) != 0;
        }

        private bool? GetTableLookFlag(int flag) => (GetTableLookMask(_tableProperties?.TableLook) & flag) == flag;

        private void SetTableLookFlag(int flag, bool? value) {
            if (value == null) {
                return;
            }

            CheckTableProperties();
            var tableLook = _tableProperties!.TableLook;
            var mask = GetTableLookMask(tableLook);
            if (tableLook == null) {
                tableLook = new TableLook();
                _tableProperties.Append(tableLook);
            }
            if (value.Value) {
                mask |= flag;
            } else {
                mask &= ~flag;
            }
            SetTableLookMask(tableLook, mask);
        }

        /// <summary>Reads Word's native defaults; expanded attributes collectively replace packed val and default on.</summary>
        internal static int GetTableLookMask(TableLook? tableLook) {
            if (tableLook == null) return NativeDefaultTableLook;
            if (!HasExpandedTableLookFlags(tableLook)) {
                string? value = tableLook.Val?.Value;
                if (string.IsNullOrWhiteSpace(value)) return NativeDefaultTableLook;
                return int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out int parsed) ? parsed : 0;
            }

            // Native Word uses on for omitted expanded attributes. Their presence makes
            // the packed value irrelevant, including when that packed value is zero.
            int mask = ExpandedDefaultTableLook;
            ApplyExpandedTableLookFlag(tableLook.FirstRow, TableLookFirstRow, ref mask);
            ApplyExpandedTableLookFlag(tableLook.LastRow, TableLookLastRow, ref mask);
            ApplyExpandedTableLookFlag(tableLook.FirstColumn, TableLookFirstColumn, ref mask);
            ApplyExpandedTableLookFlag(tableLook.LastColumn, TableLookLastColumn, ref mask);
            ApplyExpandedTableLookFlag(tableLook.NoHorizontalBand, TableLookNoHorizontalBand, ref mask);
            ApplyExpandedTableLookFlag(tableLook.NoVerticalBand, TableLookNoVerticalBand, ref mask);
            return mask;
        }

        /// <summary>Whether expanded flags, rather than the legacy packed mask, are authoritative.</summary>
        internal static bool HasExpandedTableLookFlags(TableLook tableLook) =>
            tableLook.FirstRow != null || tableLook.LastRow != null || tableLook.FirstColumn != null ||
            tableLook.LastColumn != null || tableLook.NoHorizontalBand != null || tableLook.NoVerticalBand != null;

        private static void ApplyExpandedTableLookFlag(OnOffValue? value, int flag, ref int mask) {
            if (value == null) {
                return;
            }

            if (value.Value) {
                mask |= flag;
            } else {
                mask &= ~flag;
            }
        }

        private static void SetTableLookMask(TableLook tableLook, int mask) {
            tableLook.Val = mask.ToString("X4", CultureInfo.InvariantCulture);
            tableLook.FirstRow = null;
            tableLook.LastRow = null;
            tableLook.FirstColumn = null;
            tableLook.LastColumn = null;
            tableLook.NoHorizontalBand = null;
            tableLook.NoVerticalBand = null;
        }
    }
}
