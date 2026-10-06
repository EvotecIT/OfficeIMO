using System.IO;

namespace OfficeIMO.Excel {
    /// <summary>Prevents package recovery from replacing a part whose declared length is invalid.</summary>
    internal static class ExcelPackagePartLengthFailure {
        private const string DataKey = "OfficeIMO.Excel.PackagePartLengthMismatch";

        internal static InvalidDataException Create(string message) {
            var exception = new InvalidDataException(message);
            exception.Data[DataKey] = true;
            return exception;
        }

        internal static bool Is(Exception exception) =>
            exception is InvalidDataException && exception.Data.Contains(DataKey);
    }
}
