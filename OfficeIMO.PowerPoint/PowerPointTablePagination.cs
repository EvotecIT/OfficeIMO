using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>Configures deterministic table continuation slides.</summary>
    public sealed class PowerPointTablePaginationOptions {
        private double _minimumRowHeightPoints = 24D;
        private int? _maximumDataRowsPerSlide;

        /// <summary>Table placement available on every generated slide.</summary>
        public PowerPointLayoutBox TableBounds { get; set; } = PowerPointLayoutBox.FromInches(0.75, 1.35, 8.5, 3.7);

        /// <summary>Slide master index used for generated pages.</summary>
        public int MasterIndex { get; set; }

        /// <summary>Slide layout index used for generated pages.</summary>
        public int LayoutIndex { get; set; }

        /// <summary>Whether the first generated table includes column headers.</summary>
        public bool IncludeHeaders { get; set; } = true;

        /// <summary>Whether continuation pages repeat column headers.</summary>
        public bool RepeatHeaders { get; set; } = true;

        /// <summary>Minimum row height used to derive page capacity from <see cref="TableBounds"/>.</summary>
        public double MinimumRowHeightPoints {
            get => _minimumRowHeightPoints;
            set {
                if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) {
                    throw new ArgumentOutOfRangeException(nameof(MinimumRowHeightPoints),
                        "Minimum row height must be a finite positive number.");
                }
                _minimumRowHeightPoints = value;
            }
        }

        /// <summary>Optional explicit data-row capacity. Null derives capacity from the table height.</summary>
        public int? MaximumDataRowsPerSlide {
            get => _maximumDataRowsPerSlide;
            set {
                if (value <= 0) {
                    throw new ArgumentOutOfRangeException(nameof(MaximumDataRowsPerSlide),
                        "Maximum rows per slide must be positive.");
                }
                _maximumDataRowsPerSlide = value;
            }
        }

        /// <summary>Called after each page slide is created and before its table is added.</summary>
        public Action<PowerPointSlide, PowerPointTablePageContext>? ConfigureSlide { get; set; }

        /// <summary>Called after each page table is populated and row heights are applied.</summary>
        public Action<PowerPointTable, PowerPointTablePageContext>? ConfigureTable { get; set; }
    }

    /// <summary>Context supplied while a paginated table page is built.</summary>
    public sealed class PowerPointTablePageContext {
        internal PowerPointTablePageContext(int pageIndex, int pageCount, int firstDataRowIndex,
            int dataRowCount, bool includesHeaders) {
            PageIndex = pageIndex;
            PageCount = pageCount;
            FirstDataRowIndex = firstDataRowIndex;
            DataRowCount = dataRowCount;
            IncludesHeaders = includesHeaders;
        }

        /// <summary>Zero-based page index.</summary>
        public int PageIndex { get; }

        /// <summary>Total page count.</summary>
        public int PageCount { get; }

        /// <summary>Zero-based source data-row index of the first row on this page.</summary>
        public int FirstDataRowIndex { get; }

        /// <summary>Number of source data rows on this page.</summary>
        public int DataRowCount { get; }

        /// <summary>Whether this page repeats the table headers.</summary>
        public bool IncludesHeaders { get; }

        /// <summary>Whether this page follows an earlier page.</summary>
        public bool IsContinuation => PageIndex > 0;
    }

    /// <summary>Slides and native editable tables produced by one pagination operation.</summary>
    public sealed class PowerPointPaginatedTableResult {
        internal PowerPointPaginatedTableResult(int sourceRowCount, IList<PowerPointSlide> slides,
            IList<PowerPointTable> tables) {
            SourceRowCount = sourceRowCount;
            Slides = new ReadOnlyCollection<PowerPointSlide>(new List<PowerPointSlide>(slides));
            Tables = new ReadOnlyCollection<PowerPointTable>(new List<PowerPointTable>(tables));
        }

        /// <summary>Number of input data rows represented by the result.</summary>
        public int SourceRowCount { get; }

        /// <summary>Generated continuation slides.</summary>
        public IReadOnlyList<PowerPointSlide> Slides { get; }

        /// <summary>Generated native PowerPoint tables in page order.</summary>
        public IReadOnlyList<PowerPointTable> Tables { get; }

        /// <summary>Number of generated pages.</summary>
        public int PageCount => Slides.Count;
    }

    public sealed partial class PowerPointPresentation {
        /// <summary>
        ///     Adds as many slides as required to represent every data row in native PowerPoint tables.
        ///     Capacity is measured from the table box and minimum row height, and headers repeat by default.
        /// </summary>
        public PowerPointPaginatedTableResult AddTableSlides<T>(IEnumerable<T> data,
            IEnumerable<PowerPointTableColumn<T>> columns, PowerPointTablePaginationOptions? options = null) {
            if (data == null) throw new ArgumentNullException(nameof(data));
            if (columns == null) throw new ArgumentNullException(nameof(columns));

            List<T> rows = data.ToList();
            if (rows.Count == 0) {
                throw new ArgumentException("Paginated table data cannot be empty.", nameof(data));
            }
            List<PowerPointTableColumn<T>> bindings = columns.ToList();
            if (bindings.Count == 0) {
                throw new ArgumentException("Paginated tables require at least one column.", nameof(columns));
            }

            PowerPointTablePaginationOptions resolved = options ?? new PowerPointTablePaginationOptions();
            ValidateTableBounds(resolved.TableBounds);
            List<PageSlice> pages = CreatePageSlices(rows.Count, resolved);
            var slides = new List<PowerPointSlide>(pages.Count);
            var tables = new List<PowerPointTable>(pages.Count);

            for (int pageIndex = 0; pageIndex < pages.Count; pageIndex++) {
                PageSlice page = pages[pageIndex];
                bool includeHeaders = pageIndex == 0 ? resolved.IncludeHeaders : resolved.RepeatHeaders;
                var context = new PowerPointTablePageContext(pageIndex, pages.Count, page.Offset, page.Count,
                    includeHeaders);
                PowerPointSlide slide = AddSlide(resolved.MasterIndex, resolved.LayoutIndex);
                resolved.ConfigureSlide?.Invoke(slide, context);

                int physicalRows = page.Count + (includeHeaders ? 1 : 0);
                double tableHeightPoints = Math.Min(resolved.TableBounds.HeightPoints,
                    physicalRows * resolved.MinimumRowHeightPoints);
                PowerPointLayoutBox tableBounds = new PowerPointLayoutBox(resolved.TableBounds.Left,
                    resolved.TableBounds.Top, resolved.TableBounds.Width,
                    PowerPointUnits.FromPoints(tableHeightPoints));
                PowerPointTable table = slide.AddTable(rows.GetRange(page.Offset, page.Count), bindings,
                    tableBounds, includeHeaders);
                ApplyRowHeights(table, resolved.MinimumRowHeightPoints);
                resolved.ConfigureTable?.Invoke(table, context);
                slides.Add(slide);
                tables.Add(table);
            }

            return new PowerPointPaginatedTableResult(rows.Count, slides, tables);
        }

        private static List<PageSlice> CreatePageSlices(int rowCount, PowerPointTablePaginationOptions options) {
            int physicalCapacity = Math.Max(1,
                (int)Math.Floor(options.TableBounds.HeightPoints / options.MinimumRowHeightPoints));
            int offset = 0;
            int pageIndex = 0;
            var pages = new List<PageSlice>();
            while (offset < rowCount) {
                bool includeHeaders = pageIndex == 0 ? options.IncludeHeaders : options.RepeatHeaders;
                int derivedDataCapacity = physicalCapacity - (includeHeaders ? 1 : 0);
                if (derivedDataCapacity <= 0) {
                    throw new InvalidOperationException(
                        "The table bounds must fit at least one data row in addition to the repeated header.");
                }

                int capacity = options.MaximumDataRowsPerSlide.HasValue
                    ? Math.Min(derivedDataCapacity, options.MaximumDataRowsPerSlide.Value)
                    : derivedDataCapacity;
                int count = Math.Min(capacity, rowCount - offset);
                pages.Add(new PageSlice(offset, count));
                offset += count;
                pageIndex++;
            }
            return pages;
        }

        private static void ApplyRowHeights(PowerPointTable table, double rowHeightPoints) {
            IReadOnlyList<PowerPointTableRow> rows = table.RowItems;
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
                rows[rowIndex].HeightPoints = rowHeightPoints;
            }
        }

        private static void ValidateTableBounds(PowerPointLayoutBox bounds) {
            if (bounds.Left < 0L || bounds.Top < 0L || bounds.Width <= 0L || bounds.Height <= 0L) {
                throw new ArgumentOutOfRangeException(nameof(bounds),
                    "Table bounds must use non-negative coordinates and positive dimensions.");
            }
        }

        private readonly struct PageSlice {
            internal PageSlice(int offset, int count) {
                Offset = offset;
                Count = count;
            }

            internal int Offset { get; }
            internal int Count { get; }
        }
    }
}
