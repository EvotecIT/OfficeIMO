using System.Runtime.CompilerServices;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        // Named styles can replace cell formats without changing their indexes or count.
        // The revision belongs to the shared format table, including uncached sheet wrappers.
        private static readonly ConditionalWeakTable<CellFormats, CellFormatCacheRevision> CellFormatCacheRevisions = new();
        private CellFormatCacheRevision? _cellValueStyleRevision;
        private long _cellValueStyleRevisionNumber;

        private sealed class CellFormatCacheRevision {
            internal long Number;
        }

        /// <summary>Reuses automatic number formats only while their shared base table is unchanged.</summary>
        private uint GetAutomaticNumberFormatStyleIndex(uint baseStyleIndex, bool duration) {
            RefreshAutomaticCellFormatCaches();
            if (baseStyleIndex == 0U) {
                uint? cached = duration ? _cellValueDefaultDurationStyleIndex : _cellValueDefaultDateStyleIndex;
                if (cached.HasValue) return cached.Value;
            } else {
                Dictionary<uint, uint>? indexes = duration ? _cellValueDurationStyleIndexes : _cellValueDateStyleIndexes;
                if (indexes != null && indexes.TryGetValue(baseStyleIndex, out uint cached)) return cached;
            }

            uint index = GetOrCreateBuiltInNumberFormatStyleIndex(baseStyleIndex, duration ? 46U : 14U);
            // Creating the first automatic format can also create the shared format table.
            RefreshAutomaticCellFormatCaches();
            if (duration) {
                if (baseStyleIndex == 0U) _cellValueDefaultDurationStyleIndex = index;
                else (_cellValueDurationStyleIndexes ??= new Dictionary<uint, uint>())[baseStyleIndex] = index;
            } else {
                if (baseStyleIndex == 0U) _cellValueDefaultDateStyleIndex = index;
                else (_cellValueDateStyleIndexes ??= new Dictionary<uint, uint>())[baseStyleIndex] = index;
            }
            return index;
        }

        private static void InvalidateAutomaticCellFormatCaches(CellFormats formats) =>
            CellFormatCacheRevisions.GetValue(formats, static _ => new CellFormatCacheRevision()).Number++;

        /// <summary>Discards automatic date/time formats after a shared format replacement or rollback.</summary>
        private void RefreshAutomaticCellFormatCaches() {
            CellFormats? formats = _excelDocument.WorkbookPartRoot.WorkbookStylesPart?.Stylesheet?.CellFormats;
            CellFormatCacheRevision? revision = formats == null ? null
                : CellFormatCacheRevisions.GetValue(formats, static _ => new CellFormatCacheRevision());
            if (ReferenceEquals(_cellValueStyleRevision, revision)
                && _cellValueStyleRevisionNumber == (revision?.Number ?? 0)) return;

            _cellValueStyleRevision = revision;
            _cellValueStyleRevisionNumber = revision?.Number ?? 0;
            _cellValueDateStyleIndexes = null;
            _cellValueDurationStyleIndexes = null;
            _cellValueDefaultDateStyleIndex = null;
            _cellValueDefaultDurationStyleIndex = null;
        }
    }
}
