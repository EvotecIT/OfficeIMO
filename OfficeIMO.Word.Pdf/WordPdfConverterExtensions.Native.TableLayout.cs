using System.Globalization;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Word.Pdf {

    public static partial class WordPdfConverterExtensions {
        private const double NativeAutoFitGridMinimumScale = 0.8D;

        private static (double Left, double Right) GetNativeTableOuterBorderInsets(TableLayout layout, PdfCore.PdfTableStyle style) {
            if (style.BorderFrame?.Border is { } perimeter) {
                return ((GetNativeFrameSide(perimeter, top: null)?.PaintThickness ?? 0D) / 2D,
                    (GetNativeFrameSide(perimeter, top: null, right: true)?.PaintThickness ?? 0D) / 2D);
            }
            double left = 0D, right = 0D;
            int columns = GetNativeTableColumnCount(layout);
            for (int row = 0; row < layout.Rows.Count; row++) {
                int column = GetNativeTableRowStartColumn(layout, row);
                foreach (WordTableCell cell in layout.Rows[row]) {
                    if (IsNativeHorizontalMergeContinuation(cell)) continue;
                    int span = GetNativeCellColumnSpan(cell);
                    if (column == 0) left = Math.Max(left, GetNativeTableCellBorderInset(style, row, column, right: false));
                    if (column + span == columns) right = Math.Max(right, GetNativeTableCellBorderInset(style, row, column, right: true));
                    column += span;
                }
            }
            return (left, right);
        }

        private static void ApplyNativeAutomaticTableFrameWidth(TableLayout layout, PdfCore.PdfTableStyle style, double? contentWidth) {
            if (!style.AutoFitUnspecifiedWidthToContent || style.PreferredWidth.HasValue ||
                style.AutoFitColumnPreferredWidthPoints == null || !contentWidth.HasValue) return;
            var (left, right) = GetNativeTableOuterBorderInsets(layout, style);
            double width = contentWidth.Value - style.LeftIndent.ToPoints() - left - right -
                (style.BorderFrame == null ? 0D : 2D * style.CellSpacing.ToPoints());
            if (!style.AutoFitColumns) {
                // A fixed automatic frame can overflow its container. Its
                // source preferences and physical cell minima remain usable.
                double minimumFrame = style.AutoFitColumnPreferredWidthPoints.Select((preferred, column) =>
                    Math.Max(preferred ?? 0D, style.ColumnMinWidths != null && column < style.ColumnMinWidths.Count
                        ? style.ColumnMinWidths[column]?.ToPoints() ?? 0D : 0D)).Sum();
                width = Math.Max(width, minimumFrame + Math.Max(0, layout.ColumnWidths.Length - 1) * style.CellSpacing.ToPoints());
            }
            if (width > .001D) style.MaxWidth = global::OfficeIMO.OfficeLength.Points(Math.Min(style.MaxWidth?.ToPoints() ?? width, width));
        }

        private static void ApplyNativeFixedPercentageColumnPreferences(TableLayout layout, PdfCore.PdfTableStyle style) {
            if (style.AutoFitColumns || style.ImportedColumnWidthFractions == null) return;
            var absolute = new bool[layout.ColumnWidths.Length];
            var cached = new bool[layout.ColumnWidths.Length];
            foreach (var cell in EnumerateNativeTableCells(layout)) {
                bool hasAbsoluteWidth = WordTableWidthNative.AbsoluteCellPoints(cell.Cell).HasValue;
                bool hasCachedWidth = !GetNativeTableCellPreferredWidthPercent(cell.Cell).HasValue;
                for (int column = cell.Column; column < cell.Column + cell.ColumnSpan && column < cached.Length; column++) {
                    absolute[column] |= hasAbsoluteWidth;
                    cached[column] |= hasCachedWidth;
                }
            }
            style.AutoFitColumnPreferredWidthPoints = layout.ColumnWidths.Select((width, column) => {
                if (!absolute[column] && !cached[column] && style.ImportedColumnWidthFractions[column].HasValue)
                    return (double?)null;
                // Separate-border cached grids include the outer and intercell
                // spacing shares. Absolute cell preferences describe the box.
                double spacingShare = !absolute[column] && style.BorderFrame != null
                    ? style.CellSpacing.ToPoints() * (cached.Length == 1 ? 2D : column == 0 || column == cached.Length - 1 ? 1.5D : 1D) : 0D;
                return (double?)Math.Max(.001D, width - spacingShare);
            }).ToList();
        }

        private static void ApplyNativeColumnWidths(WordTable table, TableLayout layout, PdfCore.PdfTableStyle style, double? contentWidth) {
            ApplyNativeAbsoluteNoWrapMinimums(layout, style);
            if (style.AutoFitColumns) {
                style.AutoFitCellPreferredWidthPoints =
                    EnumerateNativeTableCellPositions(layout)
                        .Where(cell => WordTableWidthNative.HasAbsoluteCellPreference(cell.Cell))
                        .ToDictionary(cell => (cell.Row, cell.Column), cell => Math.Max(.001D,
                            WordTableWidthNative.AbsoluteCellPoints(cell.Cell) ??
                            layout.ColumnWidths.Skip(cell.Column).Take(cell.ColumnSpan).Sum(width => (double)width)));
            }
            if (!style.AutoFitColumns && style.ImportedColumnWidthFractions == null && !style.MaxWidth.HasValue &&
                layout.ColumnWidths.Length > 0 && layout.ColumnWidths.All(width => width >= 0F)) {
                double gapWidth = Math.Max(0, layout.ColumnWidths.Length - 1) * style.CellSpacing.ToPoints();
                style.MaxWidth = global::OfficeIMO.OfficeLength.Points(Math.Max(layout.ColumnWidths.Sum(width => (double)width),
                    style.ColumnMinWidths?.ConvertAll(static length => length?.ToPoints())?.Sum(width => width ?? 0D) ?? 0D) + gapWidth);
                style.PreserveWidth = true;
                style.AllowWidthBeyondFrame = true;
            }
            if (style.ImportedColumnWidthFractions != null) {
                style.ColumnWidths = null;
                style.ColumnWidthWeights = style.ImportedColumnWidthFractions.All(value => value.HasValue)
                    ? style.ImportedColumnWidthFractions.Select(value => value!.Value).ToList() : null;
                return;
            }

            if (style.AutoFitColumns) {
                if (style.AutoFitColumnPreferredWidthPoints != null) return;
                if (layout.ColumnWidths.Length > 0 && layout.ColumnWidths.All(width => width > 0)) {
                    ApplyNativeAutoFitGridMinimums(layout, style, contentWidth);
                }
                if (layout.ColumnWidths.Length > 0 &&
                    layout.ColumnWidths.All(width => width > 0) &&
                    (!IsNativeTableAutoFitToContents(table.PropertiesNode()) || HasNativeTableCellWidthPreferences(layout)) &&
                    layout.ColumnWidths.Skip(1).Any(width => Math.Abs(width - layout.ColumnWidths[0]) > 0.01F)) {
                    style.ColumnWidthWeights = layout.ColumnWidths.Select(width => (double)width).ToList();
                }
                return;
            }

            style.ColumnWidths = (CreateNativeColumnWidthPoints(layout, style)) is { } lengths7086 ? System.Linq.Enumerable.ToList(System.Linq.Enumerable.Select(lengths7086, static length => length.HasValue ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(length.Value) : null)) : null;
        }

        /// <summary>No-wrap absolute cells can wrap words, but retain their preferred allocation before other cells shrink.</summary>
        private static void ApplyNativeAbsoluteNoWrapMinimums(TableLayout layout, PdfCore.PdfTableStyle style) {
            if (!style.AutoFitColumns) return;
            List<double?>? minimums = null;
            foreach (var cell in EnumerateNativeTableCellPositions(layout)) {
                double? width = WordTableWidthNative.AbsoluteCellPoints(cell.Cell);
                if (cell.Cell.WrapText || width is not > 0D) continue;
                minimums ??= style.ColumnMinWidths?.ConvertAll(static length => length?.ToPoints()) ?? new List<double?>();
                while (minimums.Count < layout.ColumnWidths.Length) minimums.Add(null);
                double borderWidth = style.BorderFrame?.Border == null ? 0D : 2D * (
                    GetNativeTableCellBorderInset(style, cell.Row, cell.Column, right: false) +
                    GetNativeTableCellBorderInset(style, cell.Row, cell.Column, right: true));
                ApplyNativeSpanMinimum(layout, minimums, cell.Column, cell.ColumnSpan,
                    Math.Max(0D, width.Value + borderWidth - Math.Max(0, cell.ColumnSpan - 1) * style.CellSpacing.ToPoints()));
            }
            if (minimums != null) {
                style.ColumnMinWidths = minimums.ConvertAll(static length =>
                    length.HasValue ? (OfficeLength?)OfficeLength.Points(length.Value) : null);
            }
        }

        private static void ApplyNativeAutoFitGridMinimums(TableLayout layout, PdfCore.PdfTableStyle style, double? contentWidth) {
            double tableWidth = style.MaxWidth?.ToPoints() ?? style.PreferredWidth?.ToPoints() ?? contentWidth ?? 0D;
            double gridWidth = layout.ColumnWidths.Sum(width => (double)width);
            if (tableWidth <= 0D ||
                gridWidth <= 0D ||
                double.IsNaN(tableWidth) ||
                double.IsInfinity(tableWidth)) {
                return;
            }

            double minimumScale = style.AutoFitWidthUsesContentMinimum ? 1D : NativeAutoFitGridMinimumScale;
            if (style.BorderFrame != null)
                tableWidth -= Math.Max(0, layout.ColumnWidths.Length - 1) * style.CellSpacing.ToPoints();
            List<double?> derivedMinimums = layout.ColumnWidths
                .Select(width => (double?)(tableWidth * width / gridWidth * minimumScale))
                .ToList();
            if (style.AutoFitWidthUsesContentMinimum) {
                // Only absolute cell preferences reserve the authored grid width.
                // Automatic cells share the remaining width according to content.
                var hasAbsolutePreference = new bool[derivedMinimums.Count];
                foreach (var cell in EnumerateNativeTableCells(layout)) {
                    if (!WordTableWidthNative.HasAbsoluteCellPreference(cell.Cell))
                        continue;
                    for (int column = cell.Column; column < cell.Column + cell.ColumnSpan && column < hasAbsolutePreference.Length; column++)
                        hasAbsolutePreference[column] = true;
                }
                for (int column = 0; column < derivedMinimums.Count; column++) {
                    if (!hasAbsolutePreference[column])
                        derivedMinimums[column] = null;
                }
            }
            if (style.ColumnMinWidths == null || style.ColumnMinWidths.Count == 0) {
                style.ColumnMinWidths = (derivedMinimums) is { } lengths10422 ? System.Linq.Enumerable.ToList(System.Linq.Enumerable.Select(lengths10422, static length => length.HasValue ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(length.Value) : null)) : null;
                return;
            }

            var mergedMinimums = style.ColumnMinWidths?.ConvertAll(static length => length?.ToPoints()) ?? new List<double?>();
            while (mergedMinimums.Count < derivedMinimums.Count) {
                mergedMinimums.Add(null);
            }

            for (int columnIndex = 0; columnIndex < derivedMinimums.Count; columnIndex++) {
                mergedMinimums[columnIndex] ??= derivedMinimums[columnIndex];
            }

            style.ColumnMinWidths = (mergedMinimums) is { } lengths10908 ? System.Linq.Enumerable.ToList(System.Linq.Enumerable.Select(lengths10908, static length => length.HasValue ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(length.Value) : null)) : null;
        }

        private static List<double?>? CreateNativeColumnWidthFractions(TableLayout layout) {
            int columnCount = GetNativeTableColumnCount(layout);
            if (columnCount == 0) {
                return null;
            }

            var weights = new double[columnCount];
            var hasPercentWidth = new bool[columnCount];
            bool hasAnyPercentWidth = false;
            foreach ((WordTableCell Cell, int Column, int ColumnSpan) cell in EnumerateNativeTableCells(layout).OrderBy(cell => cell.ColumnSpan)) {
                double? percent = GetNativeTableCellPreferredWidthPercent(cell.Cell);
                if (!percent.HasValue) {
                    continue;
                }

                double columnWeight = percent.Value / cell.ColumnSpan;
                if (cell.ColumnSpan > 1) {
                    // Individual cells constrain their own columns. A merged
                    // preference supplies only the still-unassigned share.
                    double assigned = 0D;
                    int missing = 0;
                    for (int column = cell.Column; column < cell.Column + cell.ColumnSpan && column < weights.Length; column++) {
                        if (hasPercentWidth[column]) assigned += weights[column];
                        else missing++;
                    }
                    if (missing == 0) continue;
                    if (percent.Value > assigned) columnWeight = (percent.Value - assigned) / missing;
                }
                for (int columnIndex = cell.Column; columnIndex < cell.Column + cell.ColumnSpan && columnIndex < weights.Length; columnIndex++) {
                    if (cell.ColumnSpan > 1 && hasPercentWidth[columnIndex]) continue;
                    if (!hasPercentWidth[columnIndex] || columnWeight > weights[columnIndex]) {
                        weights[columnIndex] = columnWeight;
                        hasPercentWidth[columnIndex] = true;
                    }
                }

                hasAnyPercentWidth = true;
            }

            if (!hasAnyPercentWidth) {
                return null;
            }

            return weights.Select((weight, column) => hasPercentWidth[column] ? (double?)weight : null).ToList();
        }

        private static List<double?>? CreateNativeColumnWidthPoints(TableLayout layout, PdfCore.PdfTableStyle style) {
            if (style.AutoFitColumns || layout.ColumnWidths.Length == 0 || !layout.ColumnWidths.All(width => width >= 0)) {
                return null;
            }

            // Core fits the complete grid once, locking columns at their
            // physical minimums before redistributing the remaining width.
            // Pre-scaling here can make a narrow column invalid before that fit.
            return layout.ColumnWidths.Select((width, column) => (double?)Math.Max(Math.Max(.001D, width),
                style.ColumnMinWidths != null && column < style.ColumnMinWidths.Count
                    ? style.ColumnMinWidths[column]?.ToPoints() ?? 0D : 0D)).ToList();
        }

        private static void ApplyNativeTableLayoutOptions(WordTable table, TableLayout layout, PdfCore.PdfTableStyle style, double? contentWidth, NativeTableStyleDefaults tableStyleDefaults) {
            OfficeXmlNode properties = table.PropertiesNode();
            if (ShouldUseNativeAutoFitTableLayout(properties, tableStyleDefaults)) {
                style.AutoFitColumns = true;
            }
            style.ImportedColumnWidthFractions = CreateNativeColumnWidthFractions(layout);

            double? cellSpacing = GetNativeTableCellSpacing(properties.WChild("tblCellSpacing")) ?? tableStyleDefaults.CellSpacing;
            if (cellSpacing.HasValue) {
                style.CellSpacing = global::OfficeIMO.OfficeLength.Points(cellSpacing.Value);
            }

            double? preferredWidth = GetNativeTablePreferredWidth(properties.WChild("tblW"), contentWidth) ??
                GetNativeTablePreferredWidth(tableStyleDefaults.PreferredWidth, contentWidth);
            if (preferredWidth.HasValue && !style.AutoFitColumns) {
                style.MaxWidth = global::OfficeIMO.OfficeLength.Points(preferredWidth.Value);
                style.PreserveWidth = true;
                // Fixed source grids retain their width in narrow columns and
                // beyond the page text area, including after indentation.
                style.AllowWidthBeyondFrame = true;
            } else {
                preferredWidth ??= GetNativeAutoFitGridPreferredWidth(properties, layout, contentWidth, style.CellSpacing.ToPoints());
                if (preferredWidth.HasValue) {
                    // An automatic table's preferred width can grow to fit an
                    // unbreakable cell; it is not a clipping boundary.
                    style.PreferredWidth = global::OfficeIMO.OfficeLength.Points(preferredWidth.Value);
                    OfficeXmlNode authoredWidth = properties.WChild("tblW");
                    if (!GetNativeTablePreferredWidth(authoredWidth, contentWidth).HasValue)
                        authoredWidth = tableStyleDefaults.PreferredWidth;
                    if (style.AutoFitColumns && ResolveNativeTablePreferredWidthType(authoredWidth) == "dxa") {
                        style.AllowWidthBeyondFrame = true;
                        style.AutoFitColumnPreferredWidthPoints = CreateNativeAutoFitColumnPreferences(layout);
                    }
                    // Positive spacing receives its independent perimeter and
                    // final cell-grid minimums after cell formatting is applied.
                    style.AutoFitWidthUsesContentMinimum = style.AutoFitColumns && style.CellSpacing.ToPoints() <= 0D;
                    style.PreserveWidth = true;
                }
            }

            if (style.AutoFitColumns && IsNativeTableAutoFitToContents(properties)) {
                // A valid automatic grid can contain only zero/automatic widths.
                // Its physical content minimum still determines the table frame.
                style.AutoFitWidthUsesContentMinimum = true;
                style.AutoFitUnspecifiedWidthToContent = true;
                style.PreserveWidth = true;
                // An automatic frame derives its width from cell preferences
                // and content, independently of the stored grid extent.
                if (preferredWidth.HasValue) style.PreferredWidth = null;
                style.AutoFitColumnPreferredWidthPoints = CreateNativeAutoFitColumnPreferences(layout);
            }
            if (!style.AutoFitColumns && style.ImportedColumnWidthFractions != null) {
                style.AutoFitColumnPreferredWidthPoints = CreateNativeAutoFitColumnPreferences(layout);
                foreach (var cell in EnumerateNativeTableCells(layout))
                    if (!GetNativeTableCellPreferredWidthPercent(cell.Cell).HasValue)
                        for (int column = cell.Column; column < cell.Column + cell.ColumnSpan &&
                            column < style.AutoFitColumnPreferredWidthPoints.Count; column++)
                            style.AutoFitColumnPreferredWidthPoints[column] ??= Math.Max(.001D, layout.ColumnWidths[column]);
                for (int column = 0; column < style.ImportedColumnWidthFractions.Count; column++)
                    if (!style.ImportedColumnWidthFractions[column].HasValue)
                        style.AutoFitColumnPreferredWidthPoints[column] ??= Math.Max(.001D, layout.ColumnWidths[column]);
                if (ResolveNativeTablePreferredWidthType(properties.WChild("tblW")) == "auto") {
                    style.PreferredWidth = null;
                    style.AutoFitUnspecifiedWidthToContent = true;
                    style.PreserveWidth = true;
                    style.AllowWidthBeyondFrame = true;
                }
            }

            // tblInd applies only to left-aligned tables. A floating coordinate
            // remains a placement coordinate even in the column fallback path.
            double? leftIndent = GetNativeTableHorizontalPositionIndent(properties.WChild("tblpPr")) ??
                (MapNativeTableAlignment(ResolveNativeTableAlignment(table, tableStyleDefaults)) == global::OfficeIMO.OfficeAlignment.Left
                    ? GetNativeTableLeftIndent(properties.WChild("tblInd")) ?? tableStyleDefaults.LeftIndent
                    : null);
            if (leftIndent.HasValue) {
                style.LeftIndent = global::OfficeIMO.OfficeLength.Points(leftIndent.Value);
            }

        }

        private static List<double?> CreateNativeAutoFitColumnPreferences(TableLayout layout) {
            var result = Enumerable.Repeat<double?>(null, layout.ColumnWidths.Length).ToList();
            foreach (var cell in EnumerateNativeTableCells(layout)) {
                if (!WordTableWidthNative.HasAbsoluteCellPreference(cell.Cell)) continue;
                for (int column = cell.Column; column < cell.Column + cell.ColumnSpan && column < result.Count; column++)
                    result[column] = Math.Max(.001D, layout.ColumnWidths[column]);
            }
            return result;
        }

        private static double? GetNativeAutoFitGridPreferredWidth(OfficeXmlNode properties, TableLayout layout, double? contentWidth, double cellSpacing) {
            if (!IsNativeTableAutoFitToContents(properties) ||
                layout.ColumnWidths.Length == 0 ||
                !layout.ColumnWidths.All(width => width > 0F) ||
                !HasNativeTableCellWidthPreferences(layout)) {
                return null;
            }

            double gridWidth = layout.ColumnWidths.Sum(width => (double)width) +
                Math.Max(0, layout.ColumnWidths.Length - 1) * cellSpacing;
            if (gridWidth <= 0D || double.IsNaN(gridWidth) || double.IsInfinity(gridWidth)) {
                return null;
            }

            return contentWidth.HasValue && contentWidth.Value > 0D
                ? Math.Min(gridWidth, contentWidth.Value)
                : gridWidth;
        }

        private static bool HasNativeTableCellWidthPreferences(TableLayout layout) =>
            EnumerateNativeTableCells(layout).Any(cell =>
                WordTableWidthNative.HasAbsoluteCellPreference(cell.Cell) ||
                GetNativeTableCellPreferredWidthPercent(cell.Cell).HasValue);

        private static bool ShouldUseNativeAutoFitTableLayout(OfficeXmlNode properties, NativeTableStyleDefaults tableStyleDefaults) {
            string? effectiveLayout = properties.WChild("tblLayout").WAttr("type") ?? tableStyleDefaults.Layout;
            if (effectiveLayout == "fixed") {
                return false;
            }

            if (effectiveLayout == "autofit") {
                return true;
            }

            // Preferred cell widths constrain sizing without changing DOCX's default algorithm.
            return true;
        }

        private static bool IsNativeTableAutoFitToContents(OfficeXmlNode properties) =>
            IsNativeTableAutoFitLayout(properties) &&
            ResolveNativeTablePreferredWidthType(properties.WChild("tblW")) == "auto";

        private static bool IsNativeTableAutoFitLayout(OfficeXmlNode properties) =>
            properties.WChild("tblLayout").WAttr("type") != "fixed";

        /// <summary>Resolves omitted defaults and Word's zero, nil and percentage preferred-width rules.</summary>
        private static string? ResolveNativeTablePreferredWidthType(OfficeXmlNode width) =>
            WordConverterNodes.ResolveTableWidthType(width);

        private static double? GetNativeTablePreferredWidth(OfficeXmlNode width, double? contentWidth) {
            string? type = ResolveNativeTablePreferredWidthType(width);
            if (type == "pct" && width.Exists) {
                double? percent = GetNativeTablePreferredWidthPercent(width);
                if (!percent.HasValue || !contentWidth.HasValue || contentWidth.Value <= 0D) {
                    return null;
                }

                return contentWidth.Value * percent.Value;
            }

            if (type != "dxa" || width.IsNone) {
                return null;
            }

            double? points = ConvertNativeTwipsToPoints(width.WAttr("w"));
            return points > 0D ? points : null;
        }

        private static double? GetNativeTablePreferredWidthPercent(OfficeXmlNode width) {
            return GetNativeTableWidthPercent(width.WAttr("w"));
        }

        private static double? GetNativeTableCellPreferredWidthPercent(WordTableCell cell) {
            OfficeXmlNode width = cell.PropertiesNode().WChild("tcW");
            if (WordConverterNodes.ResolveTableCellWidthType(width) != "pct") {
                return null;
            }

            return GetNativeTableWidthPercent(width.WAttr("w"));
        }

        private static double? GetNativeTableWidthPercent(string? rawWidth) => WordTableWidthNative.Percentage(rawWidth);

        private static double? GetNativeTableLeftIndent(OfficeXmlNode indentation) {
            int? width = indentation.WInt32("w");
            if (indentation.WAttr("type") != "dxa" || !width.HasValue) {
                return null;
            }

            return width.Value / 20D;
        }

        private static double? GetNativeTableHorizontalPositionIndent(OfficeXmlNode position) {
            int? positionX = position.WInt32("tblpX");
            if (!positionX.HasValue || position.WAttr("tblpXSpec") != null) {
                return null;
            }

            string? anchor = position.WAttr("horzAnchor");
            if (anchor != null &&
                anchor != "margin" &&
                anchor != "text") {
                return null;
            }

            double? indent = ConvertNativeTwipsToPoints(positionX.Value);
            return indent.HasValue && indent.Value >= 0D ? indent.Value : null;
        }

        private static double? GetNativeTableCellSpacing(OfficeXmlNode spacing) {
            // Word treats these explicit units as zero rather than falling
            // through to a positive spacing value from the table style.
            string? type = spacing.WAttr("type");
            if (type == "auto" || type == "pct") return 0D;
            if (type != "dxa") {
                return null;
            }

            return ConvertNativeTwipsToPoints(spacing.WAttr("w"));
        }

        /// <summary>An absolute spacing declaration selects separate border geometry even when its value is zero.</summary>
        private static bool UsesNativeSeparateTableCellBorders(WordTable table, PdfCore.PdfTableStyle style,
            NativeTableStyleDefaults defaults) {
            if (style.CellSpacing.ToPoints() > 0D) return true;
            OfficeXmlNode spacing = table.PropertiesNode().WChild("tblCellSpacing");
            return GetNativeTableCellSpacing(spacing).HasValue
                ? spacing.WAttr("type") == "dxa"
                : defaults.SeparateCellBorders == true;
        }

        private static void ApplyNativeTableDefaultCellMargins(WordTable table, PdfCore.PdfTableStyle style, bool preserveConfiguredFallbackPadding, NativeTableStyleDefaults tableStyleDefaults) {
            if (!preserveConfiguredFallbackPadding) {
                // Standard Word cell margins apply to each missing side. A
                // partial table declaration must not lose the horizontal inset.
                style.CellPaddingTop ??= global::OfficeIMO.OfficeLength.Points(0D);
                style.CellPaddingBottom ??= global::OfficeIMO.OfficeLength.Points(0D);
                style.CellPaddingLeft ??= global::OfficeIMO.OfficeLength.Points(5.4D);
                style.CellPaddingRight ??= global::OfficeIMO.OfficeLength.Points(5.4D);
            }
            if (tableStyleDefaults.CellPadding != null) {
                ApplyNativeResolvedTableCellPadding(style, tableStyleDefaults.CellPadding);
            }
            OfficeXmlNode margins = table.PropertiesNode().WChild("tblCellMar");
            if (margins.IsNone) {
                return;
            }

            (double? top, double? bottom, double? left, double? right) = ReadNativeTableCellMarginDefault(margins);

            if (top.HasValue) {
                style.CellPaddingTop = global::OfficeIMO.OfficeLength.Points(top.Value);
            }

            if (bottom.HasValue) {
                style.CellPaddingBottom = global::OfficeIMO.OfficeLength.Points(bottom.Value);
            }

            if (left.HasValue) {
                style.CellPaddingLeft = global::OfficeIMO.OfficeLength.Points(left.Value);
            }

            if (right.HasValue) {
                style.CellPaddingRight = global::OfficeIMO.OfficeLength.Points(right.Value);
            }
        }

        /// <summary>
        /// Reads a <c>w:tblCellMar</c>: the top and bottom widths, and the physical left and right widths, as points.
        /// </summary>
        private static (double? Top, double? Bottom, double? Left, double? Right) ReadNativeTableCellMarginDefault(OfficeXmlNode margins) {
            OfficeXmlNode left = margins.WChild("left"), right = margins.WChild("right");
            return (ConvertNativeTwipsToPoints(margins.WChild("top").WAttr("w")),
                ConvertNativeTwipsToPoints(margins.WChild("bottom").WAttr("w")),
                ParseNativeShortTwips(left.WAttr("w")),
                ParseNativeShortTwips(right.WAttr("w")));
        }

        /// <summary>A short twips measurement (a <c>w:left</c>/<c>w:right</c> cell margin width) as points.</summary>
        private static double? ParseNativeShortTwips(string? value) =>
            short.TryParse(value?.Trim(), NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out short twips)
                ? ConvertNativeTwipsToPoints(twips) : null;

        private static void ApplyNativeResolvedTableCellPadding(PdfCore.PdfTableStyle style, PdfCore.PdfCellPadding padding) {
            if (padding.Top.HasValue) {
                style.CellPaddingTop = padding.Top.Value;
            }

            if (padding.Bottom.HasValue) {
                style.CellPaddingBottom = padding.Bottom.Value;
            }

            if (padding.Left.HasValue) {
                style.CellPaddingLeft = padding.Left.Value;
            }

            if (padding.Right.HasValue) {
                style.CellPaddingRight = padding.Right.Value;
            }
        }

        private static PdfCore.PdfCellPadding? CreateNativeTableCellPadding(WordTableCell cell) {
            double? top = cell.MarginTop.HasValue ? ConvertNativeTwipsToPoints(checked((short)(cell.MarginTop.Value.ToTwips()))) : null;
            double? bottom = cell.MarginBottom.HasValue ? ConvertNativeTwipsToPoints(checked((short)(cell.MarginBottom.Value.ToTwips()))) : null;
            double? left = cell.MarginLeft.HasValue ? ConvertNativeTwipsToPoints(checked((short)(cell.MarginLeft.Value.ToTwips()))) : null;
            double? right = cell.MarginRight.HasValue ? ConvertNativeTwipsToPoints(checked((short)(cell.MarginRight.Value.ToTwips()))) : null;
            if (!top.HasValue && !bottom.HasValue && !left.HasValue && !right.HasValue) {
                return null;
            }

            return new PdfCore.PdfCellPadding {
                Top = (top) is { } points29616 ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(points29616) : null,
                Bottom = (bottom) is { } points29646 ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(points29646) : null,
                Left = (left) is { } points29677 ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(points29677) : null,
                Right = (right) is { } points29707 ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(points29707) : null
            };
        }

        private static double? ConvertNativeTwipsToPoints(string? value) {
            if (!int.TryParse(value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int twips) || twips < 0) {
                return null;
            }

            return twips / 20D;
        }

        private static double? ConvertNativeTwipsToPoints(int twips) {
            return twips < 0 ? null : twips / 20D;
        }

        private static double ConvertNativeEmusToPoints(long emus) {
            return emus <= 0 ? 0D : emus / 12700D;
        }

        private static void ApplyNativeTableRowOptions(WordTable table, PdfCore.PdfTableStyle style) {
            style.AllowRowBreakAcrossPages = table.AllowRowToBreakAcrossPages;
            List<bool?>? rowBreakPolicies = GetNativeTableRowBreakPolicies(table);
            if (rowBreakPolicies != null) {
                style.RowAllowBreakAcrossPages = rowBreakPolicies;
            }

            List<double?>? rowMinHeights = GetNativeTableRowHeights(table, exact: false);
            if (rowMinHeights != null) {
                double? uniformHeight = GetNativeUniformTableRowHeight(rowMinHeights);
                if (uniformHeight.HasValue) {
                    style.MinRowHeight = global::OfficeIMO.OfficeLength.Points(uniformHeight.Value);
                } else {
                    style.RowMinHeights = (rowMinHeights) is { } lengths31069 ? System.Linq.Enumerable.ToList(System.Linq.Enumerable.Select(lengths31069, static length => length.HasValue ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(length.Value) : null)) : null;
                }
            }

            List<double?>? fixedRowHeights = GetNativeTableRowHeights(table, exact: true);
            if (fixedRowHeights != null) {
                style.FixedRowHeights = (fixedRowHeights) is { } lengths31291 ? System.Linq.Enumerable.ToList(System.Linq.Enumerable.Select(lengths31291, static length => length.HasValue ? (global::OfficeIMO.OfficeLength?)global::OfficeIMO.OfficeLength.Points(length.Value) : null)) : null;
            }
        }

        private static List<bool?>? GetNativeTableRowBreakPolicies(WordTable table) {
            var policies = new List<bool?>(table.Rows.Count);
            bool? firstPolicy = null;
            bool hasMixedPolicies = false;
            foreach (WordTableRow row in table.Rows) {
                bool policy = row.AllowRowToBreakAcrossPages;
                policies.Add(policy);
                if (!firstPolicy.HasValue) {
                    firstPolicy = policy;
                    continue;
                }

                hasMixedPolicies |= firstPolicy.Value != policy;
            }

            return hasMixedPolicies ? policies : null;
        }

        private static List<double?>? GetNativeTableRowHeights(WordTable table, bool exact) {
            var heights = new List<double?>(table.Rows.Count);
            bool hasHeight = false;
            foreach (WordTableRow row in table.Rows) {
                OfficeXmlNode rowHeight = row.Node().WChild("trPr").WChild("trHeight");
                bool isExact = rowHeight.WAttr("hRule") == "exact";
                double? height = row.Height.HasValue && checked((int)(row.Height.Value.ToTwips())) > 0 && isExact == exact
                    ? ConvertNativeTwipsToPoints(checked((int)(row.Height.Value.ToTwips())))
                    : null;
                heights.Add(height);
                hasHeight |= height.HasValue;
            }

            return hasHeight ? heights : null;
        }

        private static double? GetNativeUniformTableRowHeight(IReadOnlyList<double?> rowHeights) {
            double? height = null;
            foreach (double? rowHeight in rowHeights) {
                if (!rowHeight.HasValue) {
                    return null;
                }

                if (!height.HasValue) {
                    height = rowHeight.Value;
                    continue;
                }

                if (System.Math.Abs(height.Value - rowHeight.Value) > 0.001D) {
                    return null;
                }
            }

            return height;
        }

        private static global::OfficeIMO.OfficeAlignment MapNativeTableAlignment(string? alignment) {
            if (alignment == "center") {
                return global::OfficeIMO.OfficeAlignment.Center;
            }

            if (alignment == "right") {
                return global::OfficeIMO.OfficeAlignment.Right;
            }

            return global::OfficeIMO.OfficeAlignment.Left;
        }

        private static string? ResolveNativeTableAlignment(WordTable table, NativeTableStyleDefaults tableStyleDefaults) =>
            ResolveNativeTablePositionAlignment(table.PropertiesNode().WChild("tblpPr")) ??
            WordConverterNodes.ReadTableAlignment(table.PropertiesNode().WChild("jc"), tableStyleDefaults.RightToLeft) ??
            tableStyleDefaults.Alignment;

        private static string? ResolveNativeTablePositionAlignment(OfficeXmlNode position) {
            string? alignment = position.WAttr("tblpXSpec");
            if (alignment == "center") {
                return "center";
            }

            if (alignment == "right" || alignment == "outside") {
                return "right";
            }

            if (alignment == "left" || alignment == "inside") {
                return "left";
            }

            return null;
        }

        private static global::OfficeIMO.OfficeAlignment GetNativeCellHorizontalAlignment(WordTableCell cell) {
            global::OfficeIMO.OfficeAlignment? alignment = null;
            foreach (WordParagraph paragraph in cell.Paragraphs) {
                if (!HasNativeCellParagraphVisibleText(paragraph)) {
                    continue;
                }

                global::OfficeIMO.OfficeAlignment paragraphAlignment = ResolveNativeColumnAlign(paragraph);
                if (alignment == null) {
                    alignment = paragraphAlignment;
                } else if (alignment.Value != paragraphAlignment) {
                    return global::OfficeIMO.OfficeAlignment.Left;
                }
            }

            return alignment ?? global::OfficeIMO.OfficeAlignment.Left;
        }

    }
}
