namespace OfficeIMO.Core.Internal {
    /// <summary>
    /// Marks an exception thrown because the input is protected with an encryption method the engine cannot open, such
    /// as encrypted ZIP entries in a package, so callers such as OfficeIMO.Reader can tell it from damaged content that
    /// surfaces as the same exception type.
    /// </summary>
    internal static class OfficeUnsupportedEncryption {
        /// <summary><see cref="Exception.Data"/> key; OfficeIMO.Reader reads the same literal.</summary>
        internal const string DataKey = "OfficeIMO.UnsupportedEncryption";

        internal static TException Mark<TException>(TException exception) where TException : Exception {
            exception.Data[DataKey] = true;
            return exception;
        }
    }
}
