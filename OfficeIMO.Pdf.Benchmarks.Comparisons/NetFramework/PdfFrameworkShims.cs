#if NETFRAMEWORK
namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// The .NET Framework build generates read inputs in process; prepared input sets are a .NET 8+ feature.
/// </summary>
internal static class PdfPreparedReadInputs {
    internal static byte[] Resolve(PdfBenchmarkProducer producer, PdfBenchmarkScenario scenario) {
        if (!string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("OFFICEIMO_PDF_READ_INPUT_ROOT"))) {
            throw new NotSupportedException("Prepared PDF read inputs need the .NET 8 or later build.");
        }
        return PdfDocumentGenerators.Generate(producer, scenario);
    }
}

internal static class KeyValuePairDeconstruction {
    internal static void Deconstruct<TKey, TValue>(this KeyValuePair<TKey, TValue> pair, out TKey key, out TValue value) {
        key = pair.Key;
        value = pair.Value;
    }
}
#endif
