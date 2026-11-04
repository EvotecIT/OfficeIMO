using OfficeIMO.Core.Internal;

namespace OfficeIMO;

public static partial class OfficeConverter {
    private static async Task<OfficeConversionReport> SaveConvertedAsync<T>(OfficeConversionResult<T> converted,
        Func<T, Task> save, T? callerOwnedValue = null) where T : class, IDisposable {
        T model = converted.Value;
        try {
            await OfficeConversionFailures.RunAsync(converted.Report, () => save(model)).ConfigureAwait(false);
            return converted.Report;
        } finally {
            if (!ReferenceEquals(model, callerOwnedValue)) model.Dispose();
        }
    }
}
