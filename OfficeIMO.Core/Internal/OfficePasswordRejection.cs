namespace OfficeIMO.Core.Internal {
    /// <summary>
    /// Marks an exception thrown because a supplied password did not verify, so callers such as OfficeIMO.Reader can tell
    /// a wrong password from damaged encrypted content that surfaces as the same exception type.
    /// </summary>
    internal static class OfficePasswordRejection {
        /// <summary><see cref="Exception.Data"/> key; OfficeIMO.Reader reads the same literal.</summary>
        internal const string DataKey = "OfficeIMO.PasswordRejected";

        internal static TException Mark<TException>(TException exception) where TException : Exception {
            exception.Data[DataKey] = true;
            return exception;
        }
    }
}
