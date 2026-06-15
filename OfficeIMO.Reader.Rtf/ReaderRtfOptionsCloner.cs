namespace OfficeIMO.Reader.Rtf;

internal static class ReaderRtfOptionsCloner {
    public static ReaderRtfOptions CloneOrDefault(ReaderRtfOptions? options) {
        return options?.Clone() ?? ReaderRtfOptions.CreateOfficeIMOProfile();
    }

    public static ReaderRtfOptions? CloneNullable(ReaderRtfOptions? options) {
        return options?.Clone();
    }
}
