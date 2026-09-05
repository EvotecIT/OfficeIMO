using System.Text;

namespace OfficeIMO.AsciiDoc.Benchmarks;

internal static class AsciiDocBenchmarkValidation {
    internal static void ValidateAll() {
        foreach (string scale in AsciiDocBenchmarkCorpus.Scales) {
            AsciiDocBenchmarkFixture fixture = AsciiDocBenchmarkCorpus.Get(scale);
            AsciiDocParseResult result = Validate(fixture);
            int inputBytes = Encoding.UTF8.GetByteCount(fixture.Source);
            int outputBytes = Encoding.UTF8.GetByteCount(result.Document.ToAsciiDoc());
            Console.WriteLine(
                $"{scale,-6} input {inputBytes,10:N0} bytes | output {outputBytes,10:N0} bytes | " +
                $"blocks {result.Document.Blocks.Count,8:N0} | sections {fixture.SectionCount,6:N0} | records {fixture.RecordCount,7:N0}");
        }
    }

    internal static AsciiDocParseResult Validate(AsciiDocBenchmarkFixture fixture) {
        AsciiDocParseResult result = AsciiDocDocument.ParseResult(fixture.Source);
        if (result.HasErrors || !result.IsLossless) throw new InvalidOperationException(fixture.Scale + " did not parse losslessly.");
        string output = result.Document.ToAsciiDoc();
        if (!string.Equals(output, fixture.Source, StringComparison.Ordinal)) {
            throw new InvalidOperationException(fixture.Scale + " preserve writing changed the source.");
        }
        int headings = result.Document.BlocksOfType<AsciiDocHeading>().Count();
        int tables = result.Document.BlocksOfType<AsciiDocTableBlock>().Count();
        if (headings != fixture.SectionCount + 1 || tables != fixture.SectionCount) {
            throw new InvalidOperationException(
                $"{fixture.Scale} produced {headings} headings/{tables} tables; expected {fixture.SectionCount + 1}/{fixture.SectionCount}.");
        }
        if (!output.Contains("Record 1:", StringComparison.Ordinal)
            || !output.Contains($"Record {fixture.RecordCount}:", StringComparison.Ordinal)
            || !output.Contains("zażółć gęślą jaźń", StringComparison.Ordinal)) {
            throw new InvalidOperationException(fixture.Scale + " lost required semantic markers.");
        }
        return result;
    }
}
