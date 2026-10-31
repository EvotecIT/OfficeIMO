using OfficeIMO.Core.Internal;

namespace OfficeIMO;

public static partial class OfficeConverter {
    private static async Task<OfficeConversionReport> SaveConvertedAsync<T>(OfficeConversionResult<T> converted,
        Func<T, Task> save) where T : IDisposable {
        using T model = converted.Value;
        await OfficeConversionFailures.RunAsync(converted.Report, () => save(model)).ConfigureAwait(false);
        return converted.Report;
    }
}
