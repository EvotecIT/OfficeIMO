namespace OfficeIMO.Reader.Visio;

internal static class ReaderVisioOptionsCloner {
    public static ReaderVisioOptions CloneOrDefault(ReaderVisioOptions? options) {
        return CloneNullable(options) ?? new ReaderVisioOptions();
    }

    public static ReaderVisioOptions? CloneNullable(ReaderVisioOptions? options) {
        if (options == null) return null;

        return new ReaderVisioOptions {
            IncludeSvgPreviewAssets = options.IncludeSvgPreviewAssets,
            IncludePngPreviewAssets = options.IncludePngPreviewAssets,
            SvgOptions = options.SvgOptions,
            PngOptions = options.PngOptions
        };
    }
}
