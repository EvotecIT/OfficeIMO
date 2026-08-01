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
        var tilingPatternResourceCache =
            new Dictionary<(PdfStream Stream, PdfDictionary Resources), PdfPageTilingPatternResource?>();
        PdfDictionary? pageResources = ResolveDictionary(GetInheritedValue("Resources"));
        var activeForms = new HashSet<PdfStream>();
        var pageContentBudget = new PageContentBudget(this);
        string content = GetContentStreamContent(pageContentBudget);
        if (content.Length > 0) {
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
                        if (!IsRepresentedByDetectedTableBorder(primitive, detectedTables, size.Height)) {
                            unrepresentedCount++;
                        }
                    }
                },
                activeForms,
                retainPrimitiveData: false,
                tilingPatternResourceCache: tilingPatternResourceCache,
                textOutputBudget: textOutputBudget,
                pageContentBudget: pageContentBudget);
        }

        return (totalCount, unrepresentedCount);
    }

    private static bool IsRepresentedByDetectedTableBorder(
        PdfPageVisualPrimitive primitive,
        IReadOnlyList<StructuredTable> detectedTables,
        double pageHeight) {
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
        const double minimumVerticalPadding = 12D;
        const double maximumVerticalPadding = 24D;
        double primitiveLeft = primitive.X;
        double primitiveTop = primitive.Y;
        double primitiveRight = primitive.X + primitive.Width;
        double primitiveBottom = primitive.Y + primitive.Height;
        for (int tableIndex = 0; tableIndex < detectedTables.Count; tableIndex++) {
            StructuredTable table = detectedTables[tableIndex];
            if (table.Columns.Count == 0 || table.YTop <= table.YBottom) continue;
            double tableLeft = table.Columns[0].From;
            double tableRight = table.Columns[table.Columns.Count - 1].To;
            double tableTop = pageHeight - table.YTop;
            double tableBottom = pageHeight - table.YBottom;
            double strokeTolerance = primitive.StrokeWidth / 2D;
            double horizontalTolerance = Math.Max(
                minimumHorizontalPadding,
                Math.Min(maximumHorizontalPadding, (tableRight - tableLeft) / table.Columns.Count * 0.75D));
            double verticalTolerance = Math.Max(
                minimumVerticalPadding,
                Math.Min(maximumVerticalPadding, (tableBottom - tableTop) / Math.Max(1, table.Rows.Count - 1) / 2D));
            horizontalTolerance = Math.Max(horizontalTolerance, strokeTolerance);
            verticalTolerance = Math.Max(verticalTolerance, strokeTolerance);
            bool horizontalMatch = primitiveLeft >= tableLeft - horizontalTolerance &&
                primitiveRight <= tableRight + horizontalTolerance;
            bool topDownVerticalMatch = primitiveTop >= tableTop - verticalTolerance &&
                primitiveBottom <= tableBottom + verticalTolerance;
            if (!horizontalMatch || !topDownVerticalMatch) continue;

            double averageColumnWidth = Math.Max(1D, (tableRight - tableLeft) / table.Columns.Count);
            double averageRowHeight = Math.Max(1D, (tableBottom - tableTop) / Math.Max(1, table.Rows.Count - 1));
            double lineTolerance = Math.Max(2D, primitive.StrokeWidth * 2D);
            bool horizontalBorder = primitive.Height <= lineTolerance &&
                primitive.Width >= averageColumnWidth * 0.5D;
            bool verticalBorder = primitive.Width <= lineTolerance &&
                primitive.Height >= averageRowHeight * 0.5D &&
                IsNearDetectedColumnBoundary(primitiveLeft, table.Columns, horizontalTolerance);
            bool cellBorderRectangle = primitive.Kind == PdfPageVisualPrimitiveKind.Rectangle &&
                primitive.Width >= averageColumnWidth * 0.5D &&
                primitive.Height <= averageRowHeight * 2.5D &&
                ((IsNearDetectedColumnBoundary(primitiveLeft, table.Columns, horizontalTolerance) &&
                  IsNearDetectedColumnBoundary(primitiveRight, table.Columns, horizontalTolerance)) ||
                 (primitiveLeft <= tableLeft + horizontalTolerance &&
                  primitiveRight >= tableRight - horizontalTolerance));
            if (horizontalBorder || verticalBorder || cellBorderRectangle) return true;
        }
        return false;
    }

    private static bool IsNearDetectedColumnBoundary(
        double x,
        List<StructuredTableColumn> columns,
        double tolerance) {
        for (int index = 0; index < columns.Count; index++) {
            if (Math.Abs(x - columns[index].From) <= tolerance ||
                Math.Abs(x - columns[index].To) <= tolerance) return true;
        }
        return false;
    }
}
