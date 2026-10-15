namespace OfficeIMO.Reader.FormatInternals;

/// <summary>Keeps a native inspection quota terminal across Reader routing and aggregate operations.</summary>
internal static class NativeReaderAdmission {
    internal static T Load<T>(Func<T> load) {
        try { return load(); }
        catch (OfficeInspectionResourceLimitException exception) {
            throw TranslateLimit(exception);
        }
    }

    internal static ReaderResourceLimitException TranslateLimit(OfficeInspectionResourceLimitException exception) =>
        new ReaderResourceLimitException(exception.Owner + "." + exception.LimitName,
            exception.Maximum, exception);
}
