using OfficeIMO.Benchmarks;

namespace OfficeIMO.Excel.Benchmarks;

/// <summary>
/// Runs a lane's output check during setup and records its result in the benchmark validation log. The check must
/// compare the lane's output with the expected data; it fails by throwing, and the exception is rethrown after the
/// failure is recorded so the lane is never timed.
/// </summary>
internal static class ExcelOutputValidation {
    /// <summary>Checks one library's output. <paramref name="caseName"/> is the declared case, or empty for all of them.</summary>
    internal static void Check(string library, string caseName, string detail, Action validate) =>
        Check([library], caseName, detail, validate);

    /// <summary>Checks several libraries' outputs together, such as a comparison against one shared expected result.</summary>
    internal static void Check(string[] libraries, string caseName, string detail, Action validate) {
        try {
            validate();
        } catch (Exception exception) {
            foreach (string library in libraries) {
                BenchmarkOutputValidationLog.Failed(library, detail + ": " + exception.Message, caseName);
            }
            throw;
        }
        foreach (string library in libraries) {
            BenchmarkOutputValidationLog.Passed(library, detail, caseName);
        }
    }
}
