namespace OfficeIMO.Excel {
    /// <summary>Where <see cref="ExcelSheet.LinkByHeader"/> points each cell: a worksheet in this workbook or a URL.</summary>
    public sealed class ExcelLinkTarget {
        private ExcelLinkTarget(bool isUrl, Func<string, string>? map, string targetCell) {
            IsUrl = isUrl;
            Map = map;
            TargetCell = targetCell;
        }

        internal bool IsUrl { get; }
        internal Func<string, string>? Map { get; }
        internal string TargetCell { get; }

        /// <summary>Links each cell to a worksheet; by default the sheet named by the cell text.</summary>
        /// <param name="sheetForCellText">Maps the cell text to a sheet name; a blank result skips the cell.</param>
        /// <param name="targetCell">Destination cell on the target sheet.</param>
        public static ExcelLinkTarget Sheet(Func<string, string>? sheetForCellText = null, string targetCell = "A1") {
            if (string.IsNullOrWhiteSpace(targetCell)) throw new ArgumentNullException(nameof(targetCell));
            return new ExcelLinkTarget(isUrl: false, sheetForCellText, targetCell);
        }

        /// <summary>Links each cell to the URL built from its text.</summary>
        /// <param name="urlForCellText">Maps the cell text to a URL; a blank result skips the cell.</param>
        public static ExcelLinkTarget Url(Func<string, string> urlForCellText) {
            if (urlForCellText == null) throw new ArgumentNullException(nameof(urlForCellText));
            return new ExcelLinkTarget(isUrl: true, urlForCellText, "A1");
        }
    }

    /// <summary>Scope and display settings for <see cref="ExcelSheet.LinkByHeader"/>.</summary>
    /// <remarks>
    /// Without <see cref="Range"/> or <see cref="Table"/> the header is found in the worksheet header row and
    /// rows <see cref="FirstRow"/> to <see cref="LastRow"/> are linked. A range uses its first row as the header
    /// row; a table uses its column names and data rows.
    /// </remarks>
    public sealed class ExcelLinkByHeaderOptions {
        /// <summary>A1 range whose first row holds the header; its remaining rows are linked.</summary>
        public string? Range { get; set; }

        /// <summary>Name of the table whose column is linked.</summary>
        public string? Table { get; set; }

        /// <summary>First worksheet row to link when no range or table is given.</summary>
        public int FirstRow { get; set; } = 2;

        /// <summary>Last worksheet row to link when no range or table is given; null uses the bottom of the used range.</summary>
        public int? LastRow { get; set; }

        /// <summary>Display text for a cell; by default the cell text for sheet links and an RFC label or host for URLs.</summary>
        public Func<string, string>? Display { get; set; }

        /// <summary>Whether links get hyperlink styling (blue, underlined).</summary>
        public bool Styled { get; set; } = true;
    }

    /// <summary>Outcome of <see cref="ExcelSheet.LinkByHeader"/>.</summary>
    public sealed class ExcelLinkByHeaderResult {
        internal ExcelLinkByHeaderResult(string? failure, int column, int linkedCount, IReadOnlyList<string> skippedCells) {
            Failure = failure;
            Column = column;
            LinkedCount = linkedCount;
            SkippedCells = skippedCells;
        }

        /// <summary>Whether the header (and table, when requested) was found.</summary>
        public bool Succeeded => Failure == null;

        /// <summary>Why nothing was linked, such as a missing header or table; null on success.</summary>
        public string? Failure { get; }

        /// <summary>1-based worksheet column that was linked; 0 when nothing was found.</summary>
        public int Column { get; }

        /// <summary>Number of cells that received a link.</summary>
        public int LinkedCount { get; }

        /// <summary>A1 addresses of non-empty cells left unlinked because the target mapping returned no sheet or URL.</summary>
        public IReadOnlyList<string> SkippedCells { get; }

        internal static ExcelLinkByHeaderResult Failed(string failure) => new ExcelLinkByHeaderResult(failure, 0, 0, Array.Empty<string>());
    }
}
