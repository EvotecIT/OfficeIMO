namespace OfficeIMO.Pdf;

public sealed partial class PdfReadPage {
    internal (int TotalCount, int UnrepresentedCount) GetVisibleVisualPrimitiveCounts(
        IReadOnlyList<StructuredTable> detectedTables) {
        (double Width, double Height) size = GetVisualPageSize();
        int totalCount = 0;
        int unrepresentedCount = 0;
        var textOutputBudget = CreateTextOutputBudget();
        var visibilityBudget = new VisualGeometryBudget();
        var patternPaintCache = new Dictionary<PdfPageTilingPatternResource, bool>();
        var tilingPatternResourceCache = new TilingPatternResourceCache();
        PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        var activeForms = new HashSet<PdfStream>();
        var pageContentBudget = new PageContentBudget(this);
        string content = GetContentStreamContent(pageContentBudget);
        if (content.Length > 0) {
            PdfPageInvokedResourceNames invokedResources = GetRootInvokedResourceNames(content, pageResources);
            CollectVisualPrimitivesAndForms(
                content,
                pageResources,
                GetVisualPageTransform(),
                size.Width,
                size.Height,
                primitive => {
                    if (IsVisibleVisualPrimitive(
                            primitive,
                            size.Width,
                            size.Height,
                            visibilityBudget,
                            patternPaintCache)) {
                        totalCount++;
                        if (!IsRepresentedByDetectedTableBorder(primitive, detectedTables)) {
                            unrepresentedCount++;
                        }
                    }
                },
                activeForms,
                retainPrimitiveData: false,
                type3ImageVisitor: static (_, _, _) => { },
                tilingPatternResourceCache: tilingPatternResourceCache,
                textOutputBudget: textOutputBudget,
                pageContentBudget: pageContentBudget,
                invokedResourceNames: invokedResources);
        }

        return (totalCount, unrepresentedCount);
    }

    private bool IsRepresentedByDetectedTableBorder(
        PdfPageVisualPrimitive primitive,
        IReadOnlyList<StructuredTable> detectedTables) {
        bool hasFill = primitive.FillColor.HasValue ||
            primitive.FillGradient != null ||
            primitive.FillRadialGradient != null ||
            primitive.FillTilingPattern != null;
        bool hasStroke = primitive.StrokeWidth > 0D &&
            (primitive.StrokeColor.HasValue ||
             primitive.StrokeGradient != null ||
             primitive.StrokeRadialGradient != null ||
             primitive.StrokeTilingPattern != null);
        if (hasFill || !hasStroke) return false;

        const double minimumHorizontalPadding = 8D;
        const double maximumHorizontalPadding = 48D;
        int rotationDegrees = GetRotationDegrees();
        bool axesSwapped = rotationDegrees == 90 || rotationDegrees == 270;
        double minimumVerticalPadding = axesSwapped ? 32D : 12D;
        double maximumVerticalPadding = axesSwapped ? 48D : 24D;
        double primitiveLeft = primitive.X;
        double primitiveTop = primitive.Y;
        double primitiveRight = primitive.X + primitive.Width;
        double primitiveBottom = primitive.Y + primitive.Height;
        for (int tableIndex = 0; tableIndex < detectedTables.Count; tableIndex++) {
            StructuredTable table = detectedTables[tableIndex];
            if (table.Columns.Count == 0 || table.YTop <= table.YBottom) continue;
            GetVisualTableBoundaries(table, out double[] verticalBoundaries, out double[] horizontalBoundaries);
            if (verticalBoundaries.Length < 2 || horizontalBoundaries.Length < 2) continue;
            double tableLeft = verticalBoundaries[0];
            double tableRight = verticalBoundaries[verticalBoundaries.Length - 1];
            double tableTop = horizontalBoundaries[0];
            double tableBottom = horizontalBoundaries[horizontalBoundaries.Length - 1];
            double strokeTolerance = primitive.StrokeWidth / 2D;
            double horizontalTolerance = Math.Max(
                minimumHorizontalPadding,
                Math.Min(maximumHorizontalPadding, (tableRight - tableLeft) / Math.Max(1, verticalBoundaries.Length - 1) * 0.75D));
            double verticalTolerance = Math.Max(
                minimumVerticalPadding,
                Math.Min(maximumVerticalPadding, (tableBottom - tableTop) / Math.Max(1, horizontalBoundaries.Length - 1) * 0.75D));
            horizontalTolerance = Math.Max(horizontalTolerance, strokeTolerance);
            verticalTolerance = Math.Max(verticalTolerance, strokeTolerance);
            bool horizontalMatch = primitiveLeft >= tableLeft - horizontalTolerance &&
                primitiveRight <= tableRight + horizontalTolerance;
            bool topDownVerticalMatch = primitiveTop >= tableTop - verticalTolerance &&
                primitiveBottom <= tableBottom + verticalTolerance;
            if (!horizontalMatch || !topDownVerticalMatch) continue;

            double averageColumnWidth = Math.Max(1D, (tableRight - tableLeft) / Math.Max(1, verticalBoundaries.Length - 1));
            double averageRowHeight = Math.Max(1D, (tableBottom - tableTop) / Math.Max(1, horizontalBoundaries.Length - 1));
            double lineTolerance = Math.Max(2D, primitive.StrokeWidth * 2D);
            double rowBoundaryTolerance = Math.Max(
                lineTolerance,
                Math.Min(verticalTolerance, averageRowHeight * 0.25D));
            double rectangleRowBoundaryTolerance = axesSwapped
                ? Math.Max(rowBoundaryTolerance, verticalTolerance)
                : rowBoundaryTolerance;
            bool horizontalBorder = primitive.Height <= lineTolerance &&
                primitive.Width >= averageColumnWidth * 0.5D &&
                IsNearBoundary((primitiveTop + primitiveBottom) / 2D, horizontalBoundaries, rowBoundaryTolerance);
            bool verticalBorder = primitive.Width <= lineTolerance &&
                primitive.Height >= averageRowHeight * 0.5D &&
                IsNearBoundary((primitiveLeft + primitiveRight) / 2D, verticalBoundaries, horizontalTolerance);
            bool cellBorderRectangle = primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle &&
                primitive.Width >= averageColumnWidth * 0.5D &&
                primitive.Height <= (axesSwapped ? tableBottom - tableTop + verticalTolerance : averageRowHeight * 2.5D) &&
                IsNearBoundary(primitiveTop, horizontalBoundaries, rectangleRowBoundaryTolerance) &&
                IsNearBoundary(primitiveBottom, horizontalBoundaries, rectangleRowBoundaryTolerance) &&
                ((IsNearBoundary(primitiveLeft, verticalBoundaries, horizontalTolerance) &&
                  IsNearBoundary(primitiveRight, verticalBoundaries, horizontalTolerance)) ||
                 (primitiveLeft <= tableLeft + horizontalTolerance &&
                  primitiveRight >= tableRight - horizontalTolerance));
            if (horizontalBorder || verticalBorder || cellBorderRectangle) return true;
        }
        return false;
    }

    private void GetVisualTableBoundaries(
        StructuredTable table,
        out double[] verticalBoundaries,
        out double[] horizontalBoundaries) {
        var vertical = new List<double>();
        var horizontal = new List<double>();
        double tableLeft = table.Columns[0].From;
        double tableRight = table.Columns[table.Columns.Count - 1].To;
        double averageRowHeight = (table.YTop - table.YBottom) / Math.Max(1, table.Rows.Count - 1);
        double firstRowBoundary = table.YTop + averageRowHeight / 2D;
        double lastRowBoundary = table.YBottom - averageRowHeight / 2D;
        var columnBoundaries = new HashSet<double>();
        for (int index = 0; index < table.Columns.Count; index++) {
            columnBoundaries.Add(table.Columns[index].From);
            columnBoundaries.Add(table.Columns[index].To);
        }
        foreach (double x in columnBoundaries) {
            AddVisualBoundary(x, lastRowBoundary, x, firstRowBoundary, vertical, horizontal);
        }
        int rowCount = Math.Max(1, table.Rows.Count);
        for (int index = 0; index <= rowCount; index++) {
            double y = firstRowBoundary - index * averageRowHeight;
            AddVisualBoundary(tableLeft, y, tableRight, y, vertical, horizontal);
        }
        verticalBoundaries = vertical.Distinct().OrderBy(static value => value).ToArray();
        horizontalBoundaries = horizontal.Distinct().OrderBy(static value => value).ToArray();
    }

    private void AddVisualBoundary(
        double x1,
        double y1,
        double x2,
        double y2,
        List<double> vertical,
        List<double> horizontal) {
        PdfVisualBounds bounds = TransformBoundsToVisual(
            Math.Min(x1, x2),
            Math.Min(y1, y2),
            Math.Max(x1, x2),
            Math.Max(y1, y2));
        if (bounds.Width <= bounds.Height) vertical.Add((bounds.Left + bounds.Right) / 2D);
        else horizontal.Add((bounds.Top + bounds.Bottom) / 2D);
    }

    private static bool IsNearBoundary(double value, double[] boundaries, double tolerance) {
        for (int index = 0; index < boundaries.Length; index++) {
            if (Math.Abs(value - boundaries[index]) <= tolerance) return true;
        }
        return false;
    }
}
