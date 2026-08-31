using System.IO;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Marks configured reader-limit failures so package recovery does not retry them as
    /// recoverable Open XML shape errors.
    /// </summary>
    internal static class ExcelReadLimitFailure {
        private const string DataKey = "OfficeIMO.Excel.ReadLimitExceeded";

        internal static InvalidDataException Create(string message) {
            var exception = new InvalidDataException(message);
            exception.Data[DataKey] = true;
            return exception;
        }

        internal static bool Is(Exception exception) =>
            exception is InvalidDataException
            && exception.Data.Contains(DataKey);
    }
}
