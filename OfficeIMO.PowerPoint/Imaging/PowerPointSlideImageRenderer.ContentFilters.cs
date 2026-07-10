namespace OfficeIMO.PowerPoint {
    internal static partial class PowerPointSlideImageRenderer {
        private static bool ShouldIncludeShape(PowerPointShape shape, PowerPointImageExportOptions options) {
            if (shape is PowerPointGroupShape) return true;
            switch (shape.ShapeContentType) {
                case PowerPointShapeContentType.Picture:
                case PowerPointShapeContentType.Media:
                    return options.IncludePictures;
                case PowerPointShapeContentType.Table:
                    return options.IncludeTables;
                case PowerPointShapeContentType.Chart:
                    return options.IncludeCharts;
                case PowerPointShapeContentType.TextBox:
                    return options.IncludeTextBoxes;
                case PowerPointShapeContentType.AutoShape:
                case PowerPointShapeContentType.Connector:
                case PowerPointShapeContentType.SmartArt:
                case PowerPointShapeContentType.OleObject:
                case PowerPointShapeContentType.Unknown:
                    return options.IncludeAutoShapes;
                default:
                    return true;
            }
        }
    }
}
