namespace OfficeIMO.Drawing;

public static partial class OfficeShapePresets {
    private static OfficeShape? LeftUpArrow(double width, double height, bool horizontalFlip, bool verticalFlip) =>
        Polygon(width, height, horizontalFlip, verticalFlip,
            (0D, 0.42D), (0.32D, 0D), (0.32D, 0.24D), (0.66D, 0.24D),
            (0.66D, 0.74D), (0.9D, 0.74D), (0.52D, 1D), (0.14D, 0.74D),
            (0.38D, 0.74D), (0.38D, 0.5D), (0D, 0.5D));

    private static OfficeShape? LeftRightUpArrow(double width, double height, bool horizontalFlip, bool verticalFlip) =>
        Polygon(width, height, horizontalFlip, verticalFlip,
            (0.5D, 0D), (0.66D, 0.2D), (0.58D, 0.2D), (0.58D, 0.42D),
            (0.76D, 0.42D), (0.76D, 0.24D), (1D, 0.5D), (0.76D, 0.76D),
            (0.76D, 0.58D), (0.24D, 0.58D), (0.24D, 0.76D), (0D, 0.5D),
            (0.24D, 0.24D), (0.24D, 0.42D), (0.42D, 0.42D), (0.42D, 0.2D),
            (0.34D, 0.2D));

    private static OfficeShape? BentUpArrow(double width, double height, bool horizontalFlip, bool verticalFlip) =>
        Polygon(width, height, horizontalFlip, verticalFlip,
            (0D, 0.62D), (0.54D, 0.62D), (0.54D, 0.28D), (0.34D, 0.28D),
            (0.68D, 0D), (1D, 0.28D), (0.82D, 0.28D), (0.82D, 0.88D), (0D, 0.88D));

    private static OfficeShape? UTurnArrow(double width, double height, bool horizontalFlip, bool verticalFlip) =>
        Polygon(width, height, horizontalFlip, verticalFlip,
            (0.72D, 1D), (0.72D, 0.28D), (0.32D, 0.28D), (0.32D, 0.46D),
            (0D, 0.22D), (0.32D, 0D), (0.32D, 0.18D), (0.94D, 0.18D),
            (0.94D, 1D));
}
