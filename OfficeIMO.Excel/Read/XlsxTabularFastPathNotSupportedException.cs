namespace OfficeIMO.Excel {
    /// <summary>
    /// Signals that a forward-only XLSX package or worksheet needs the complete Open XML SDK path.
    /// It is an internal routing signal, not a public failure contract.
    /// </summary>
    internal sealed class XlsxTabularFastPathNotSupportedException : Exception {
        internal XlsxTabularFastPathNotSupportedException(string message)
            : base(message) { }

        internal XlsxTabularFastPathNotSupportedException(string message, Exception innerException)
            : base(message, innerException) { }
    }
}
