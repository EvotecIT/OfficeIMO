using OfficeIMO.Core.Internal;

namespace OfficeIMO.Excel {
    /// <summary>Preserves configured native read admission separately from malformed package recovery.</summary>
    internal static class ExcelReadLimitFailure {
        internal static OfficeInspectionResourceLimitException Create(string limitName, long maximum, string message) =>
            new OfficeInspectionResourceLimitException("excel", limitName, maximum, new InvalidDataException(message));

        internal static bool Is(Exception exception) =>
            exception is OfficeInspectionResourceLimitException limit && limit.Owner == "excel";

        internal static byte[] ReadInput(Func<byte[]> read, long maximum) {
            try { return read(); }
            catch (InvalidDataException exception) when (OfficeStreamReader.IsSizeLimitException(exception)) {
                throw CreateInputFailure(exception, maximum);
            }
        }

        private static OfficeInspectionResourceLimitException CreateInputFailure(InvalidDataException exception, long maximum) {
            long effectiveMaximum = OfficeStreamReader.GetExceededSizeLimit(exception) ?? maximum;
            string name = effectiveMaximum < maximum ? nameof(ExcelLoadOptions.MaxLegacyInputBytes) : nameof(ExcelReaderOptions.MaxInputBytes);
            return new OfficeInspectionResourceLimitException("excel", name, effectiveMaximum, exception);
        }

        internal static async Task<byte[]> ReadInputAsync(Func<Task<byte[]>> read, long maximum) {
            try { return await read().ConfigureAwait(false); }
            catch (InvalidDataException exception) when (OfficeStreamReader.IsSizeLimitException(exception)) {
                throw CreateInputFailure(exception, maximum);
            }
        }
    }
}
