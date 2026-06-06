namespace OfficeIMO.Reader.Pdf;

internal static class ReaderPdfOptionsCloner {
    public static ReaderPdfOptions CloneOrDefault(ReaderPdfOptions? options) {
        return options?.Clone() ?? ReaderPdfOptions.CreateOfficeIMOProfile();
    }

    public static ReaderPdfOptions? CloneNullable(ReaderPdfOptions? options) {
        return options?.Clone();
    }
}
