using System.Text;

namespace OfficeIMO.Latex.Benchmarks;

internal static class LatexBenchmarkValidation {
    internal static void ValidateAll() {
        foreach (string scale in LatexBenchmarkCorpus.Scales) {
            LatexBenchmarkFixture fixture = LatexBenchmarkCorpus.Get(scale);
            LatexParseResult result = Validate(fixture);
            int inputBytes = Encoding.UTF8.GetByteCount(fixture.Source);
            int outputBytes = Encoding.UTF8.GetByteCount(result.Document.ToLatex());
            Console.WriteLine(
                $"{scale,-6} input {inputBytes,10:N0} bytes | output {outputBytes,10:N0} bytes | " +
                $"tokens {result.Document.Tokens.Count,9:N0} | sections {result.Document.Headings.Count,6:N0} | " +
                $"records {fixture.RecordCount,7:N0}");
        }
    }

    internal static LatexParseResult Validate(LatexBenchmarkFixture fixture) {
        LatexParseResult result = LatexDocument.ParseResult(fixture.Source);
        if (result.HasErrors || !result.IsLossless) {
            throw new InvalidOperationException(fixture.Scale + " did not parse losslessly without errors.");
        }
        string output = result.Document.ToLatex();
        if (!string.Equals(output, fixture.Source, StringComparison.Ordinal)) {
            throw new InvalidOperationException(fixture.Scale + " preserve writing changed the source.");
        }
        if (result.Document.Headings.Count != fixture.SectionCount) {
            throw new InvalidOperationException(
                $"{fixture.Scale} produced {result.Document.Headings.Count} headings; expected {fixture.SectionCount}.");
        }
        if (!output.Contains("Record 1:", StringComparison.Ordinal)
            || !output.Contains($"Record {fixture.RecordCount}:", StringComparison.Ordinal)
            || !output.Contains("zażółć gęślą jaźń", StringComparison.Ordinal)) {
            throw new InvalidOperationException(fixture.Scale + " lost required semantic markers.");
        }
        return result;
    }

    internal static void InspectTokens(LatexDocument document) {
        int cursor = 0;
        foreach (LatexToken token in document.Tokens) {
            LatexSourceSpan span = token.Span;
            string text = token.Text;
            string? value = token.Value;
            if (span.Start.Offset != cursor || span.Length != text.Length ||
                !document.Source.Text.AsSpan(cursor, span.Length).SequenceEqual(text.AsSpan()) || !token.IsTerminated)
                throw new InvalidOperationException("Public token inspection lost exact source coverage or termination state.");
            if (token.Kind == LatexTokenKind.Command && !text.AsSpan(1).SequenceEqual(value.AsSpan()))
                throw new InvalidOperationException("Public token inspection changed a control sequence value.");
            cursor = span.End.Offset;
        }
        if (cursor != document.Source.Text.Length) throw new InvalidOperationException("Public token inspection lost source tail.");
    }
}
