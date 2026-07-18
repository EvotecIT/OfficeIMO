namespace OfficeIMO.Pdf;

internal static partial class PdfFormFiller {
    private sealed class FlattenWidgetState {
        internal FlattenWidgetState(int widgetObjectNumber, double x, double y, double width, double height, int appearanceObjectNumber) {
            WidgetObjectNumber = widgetObjectNumber;
            X = x;
            Y = y;
            Width = width;
            Height = height;
            AppearanceObjectNumber = appearanceObjectNumber;
        }

        internal int WidgetObjectNumber { get; }
        internal double X { get; }
        internal double Y { get; }
        internal double Width { get; }
        internal double Height { get; }
        internal int AppearanceObjectNumber { get; }
    }
}
