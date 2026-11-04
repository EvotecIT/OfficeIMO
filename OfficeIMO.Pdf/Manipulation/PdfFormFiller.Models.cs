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
        internal bool IsNotPainted => AppearanceObjectNumber == 0;

        internal static FlattenWidgetState NotPainted(int widgetObjectNumber) => new FlattenWidgetState(widgetObjectNumber, 0D, 0D, 0D, 0D, 0);
    }
}
