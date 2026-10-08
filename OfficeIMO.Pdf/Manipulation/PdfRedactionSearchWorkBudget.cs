using System.Diagnostics;
using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

internal sealed class PdfRedactionSearchWorkBudget {
    private const long MaximumWorkUnits = 20_000_000L;
    private static readonly TimeSpan MaximumElapsed = TimeSpan.FromSeconds(30);
    private readonly Stopwatch _elapsed = Stopwatch.StartNew();
    private readonly string _phase;
    private long _remaining = MaximumWorkUnits;

    internal PdfRedactionSearchWorkBudget(string phase) => _phase = phase;

    internal void Charge(long work) {
        if (work < 0 || work > _remaining || _elapsed.Elapsed >= MaximumElapsed)
            throw new InvalidDataException("PDF redaction " + _phase + " exceeds its logical text work limit.");
        _remaining -= work;
    }

    internal void ChargeTextScan(string text, string criterion) =>
        Charge((long)text.Length + criterion.Length + 1L);

    internal bool IsMatch(Regex regex, string text) {
        ChargeTextScan(text, regex.ToString());
        Regex bounded = BoundRegex(regex);
        bool matched;
        try {
            matched = bounded.IsMatch(text);
        } catch (RegexMatchTimeoutException) when (_elapsed.Elapsed >= MaximumElapsed) {
            Charge(0L);
            throw;
        }
        Charge(0L);
        return matched;
    }

    internal IEnumerable<Match> Matches(Regex regex, string text) {
        ChargeTextScan(text, regex.ToString());
        Regex bounded = BoundRegex(regex);
        for (Match match = bounded.Match(text); match.Success; match = match.NextMatch()) {
            Charge(1L);
            yield return match;
        }
        Charge(0L);
    }

    private Regex BoundRegex(Regex regex) {
        TimeSpan remaining = MaximumElapsed - _elapsed.Elapsed;
        if (remaining <= TimeSpan.Zero) Charge(0L);
        return regex.MatchTimeout <= TimeSpan.Zero || regex.MatchTimeout > remaining
            ? new Regex(regex.ToString(), regex.Options, remaining)
            : regex;
    }
}
