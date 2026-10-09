namespace OfficeIMO.Word.LegacyDoc.Model {
    internal static partial class LegacyDocParagraphFormattingReader {
        private const ushort SprmTInsert = 0x7621;
        private const ushort SprmTDelete = 0x5622;
        private const ushort SprmTDxaCol = 0x7623;
        private const ushort SprmTDxaLeft = 0x9601;
        private const ushort SprmTDxaGapHalf = 0x9602;

        private static bool TryChangeTableCellCount(byte[] bytes, int offset, int end, ushort sprm,
            ref IReadOnlyList<int>? widths,
            ref IReadOnlyList<LegacyDocTableCellHorizontalMerge>? horizontalMerges,
            ref IReadOnlyList<LegacyDocTableCellVerticalMerge>? verticalMerges,
            ref IReadOnlyList<LegacyDocTableCellVerticalAlignment>? verticalAlignments,
            ref IReadOnlyList<LegacyDocTableCellTextDirection>? textDirections,
            ref IReadOnlyList<bool>? fitTexts, ref IReadOnlyList<bool>? noWraps, ref IReadOnlyList<bool>? hideMarks,
            ref IReadOnlyList<LegacyDocTableCellMargins>? margins,
            ref IReadOnlyList<LegacyDocTableCellShading>? shadings,
            ref IReadOnlyList<LegacyDocTableCellBorders>? borders) {
            int length = sprm == SprmTInsert ? 6 : 4;
            if (end - offset < length) return false;
            int previousCount = widths?.Count ?? 0;
            int first = bytes[offset + 2];
            int removed = 0;
            int inserted = 0;
            int width = 0;
            if (sprm == SprmTInsert) {
                inserted = bytes[offset + 3];
                width = ReadInt16(bytes, offset + 4);
                if (inserted == 0 || width < 0 || first > previousCount || previousCount + inserted > 63)
                    return false;
            } else {
                int limit = bytes[offset + 3];
                if (first >= previousCount || first > limit || limit > previousCount || previousCount - (limit - first) == 0)
                    return false;
                if (first == limit) return true;
                removed = limit - first;
            }
            int[] changedWidths = ChangeDefinedCells(widths ?? Array.Empty<int>(), previousCount, first, removed, inserted, width);
            if (changedWidths.Sum() > 31680) return false;
            widths = changedWidths;
            horizontalMerges = ChangeCells(horizontalMerges, previousCount, first, removed, inserted, default);
            verticalMerges = ChangeCells(verticalMerges, previousCount, first, removed, inserted, default);
            verticalAlignments = ChangeCells(verticalAlignments, previousCount, first, removed, inserted, default);
            textDirections = ChangeCells(textDirections, previousCount, first, removed, inserted, default);
            fitTexts = ChangeCells(fitTexts, previousCount, first, removed, inserted, false);
            noWraps = ChangeCells(noWraps, previousCount, first, removed, inserted, false);
            hideMarks = ChangeCells(hideMarks, previousCount, first, removed, inserted, false);
            margins = ChangeCells(margins, previousCount, first, removed, inserted, default);
            shadings = ChangeCells(shadings, previousCount, first, removed, inserted, default);
            borders = ChangeCells(borders, previousCount, first, removed, inserted, default);
            return true;
        }

        private static T[]? ChangeCells<T>(IReadOnlyList<T>? source, int previousCount,
            int first, int removed, int inserted, T insertedValue) {
            // An absent exception array stays absent so inserted cells retain table-style defaults.
            if (source == null || source.Count == 0) return null;
            return ChangeDefinedCells(source, previousCount, first, removed, inserted, insertedValue);
        }

        private static T[] ChangeDefinedCells<T>(IReadOnlyList<T> source, int previousCount,
            int first, int removed, int inserted, T insertedValue) {
            var result = new T[previousCount - removed + inserted];
            for (int index = 0; index < result.Length; index++) {
                if (index >= first && index < first + inserted) {
                    result[index] = insertedValue;
                    continue;
                }
                int original = index < first ? index : index - inserted + removed;
                result[index] = original < source.Count ? source[original] : default!;
            }
            return result;
        }

        private static bool TryChangeTableColumnWidths(byte[] bytes, int offset, int end,
            ref IReadOnlyList<int>? widths) {
            if (end - offset < 6 || widths == null) return false;
            int first = bytes[offset + 2];
            int limit = bytes[offset + 3];
            int width = ReadInt16(bytes, offset + 4);
            if (first >= widths.Count || first > limit || limit > widths.Count || width < 0) return false;
            if (first == limit) return true;
            int[] result = widths.ToArray();
            for (int cell = first; cell < limit; cell++) result[cell] = width;
            if (result.Sum() > 31680) return false;
            widths = result;
            return true;
        }
    }
}
