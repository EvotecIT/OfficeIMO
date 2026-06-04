using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Globalization;

namespace OfficeIMO.Word {
    public partial class WordTable {
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

        private bool? GetTableLookFlag(int flag, Func<TableLook, OnOffValue?> expandedAttribute) {
            var tableLook = _tableProperties?.TableLook;
            if (tableLook == null) {
                return null;
            }

            var attributeValue = expandedAttribute(tableLook);
            if (attributeValue != null) {
                return attributeValue.Value;
            }

            return (GetTableLookMask(tableLook) & flag) == flag;
        }

        private void SetTableLookFlag(int flag, bool? value) {
            if (value == null) {
                return;
            }

            CheckTableProperties();
            var tableLook = _tableProperties!.TableLook;
            if (tableLook == null) {
                tableLook = CreateTableLook(firstRow: false, lastRow: false, firstColumn: false, lastColumn: false, noHorizontalBand: false, noVerticalBand: false);
                _tableProperties.Append(tableLook);
            }

            var mask = GetTableLookMask(tableLook);
            if (value.Value) {
                mask |= flag;
            } else {
                mask &= ~flag;
            }
            SetTableLookMask(tableLook, mask);
        }

        private static int GetTableLookMask(TableLook tableLook) {
            var mask = 0;
            var value = tableLook.Val?.Value;
            if (!string.IsNullOrWhiteSpace(value)
                && int.TryParse(value, NumberStyles.HexNumber, CultureInfo.InvariantCulture, out var parsed)) {
                mask = parsed;
            }

            if (tableLook.FirstRow?.Value == true) mask |= TableLookFirstRow;
            if (tableLook.LastRow?.Value == true) mask |= TableLookLastRow;
            if (tableLook.FirstColumn?.Value == true) mask |= TableLookFirstColumn;
            if (tableLook.LastColumn?.Value == true) mask |= TableLookLastColumn;
            if (tableLook.NoHorizontalBand?.Value == true) mask |= TableLookNoHorizontalBand;
            if (tableLook.NoVerticalBand?.Value == true) mask |= TableLookNoVerticalBand;

            return mask;
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
