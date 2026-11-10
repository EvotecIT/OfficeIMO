using System.Globalization;

namespace OfficeIMO.Yaml.Tests.Conformance;

/// <summary>Checks the harness runs against suite cases and corpus files.</summary>
public enum YamlConformanceCheck {
    /// <summary>A valid case produces exactly the events in <c>test.event</c>.</summary>
    Events,
    /// <summary>A valid case produces the content of <c>test.event</c>: kinds, values, tags, anchors and aliases, ignoring presentation.</summary>
    EventContent,
    /// <summary>An invalid case is rejected.</summary>
    ErrorCase,
    /// <summary>A valid case loads to values equal to <c>in.json</c>.</summary>
    Json,
    /// <summary>A corpus file parses without error.</summary>
    Accepts,
    /// <summary>Loading and saving without edits returns the input bytes.</summary>
    RoundTrip,
    /// <summary>Loading and saving without edits keeps the content: re-parsing the output gives the same values, tags and anchors.</summary>
    RoundTripContent
}

/// <summary>Outcome of one check.</summary>
public enum YamlCheckOutcome {
    /// <summary>The engine met the expectation.</summary>
    Pass,
    /// <summary>The engine produced a wrong result, or accepted invalid input.</summary>
    Fail,
    /// <summary>The engine threw or timed out on input it should handle.</summary>
    Error,
    /// <summary>The check does not apply to this input or engine.</summary>
    Skipped
}

/// <summary>Result of one check on one suite case or corpus file.</summary>
public sealed class YamlCheckResult {
    /// <summary>Creates a result.</summary>
    public YamlCheckResult(string subject, YamlConformanceCheck check, YamlCheckOutcome outcome, string? detail = null) {
        Subject = subject;
        Check = check;
        Outcome = outcome;
        Detail = detail;
    }

    /// <summary>Suite case identifier or corpus path.</summary>
    public string Subject { get; }

    /// <summary>Check that ran.</summary>
    public YamlConformanceCheck Check { get; }

    /// <summary>Outcome.</summary>
    public YamlCheckOutcome Outcome { get; }

    /// <summary>Short diagnostic for failures and errors.</summary>
    public string? Detail { get; }

    /// <summary>Returns a one-line summary.</summary>
    public override string ToString() => Subject + " " + Check + ": " + Outcome + (Detail == null ? string.Empty : " (" + Detail + ")");
}

/// <summary>Runs the conformance checks for one engine.</summary>
public sealed partial class YamlConformanceChecks {
    private readonly IYamlConformanceEngine _engine;
    private readonly TimeSpan? _timeout;

    /// <summary>
    /// Creates checks for an engine. With a timeout, each engine call runs on a worker thread and a call that does not
    /// finish in time is reported as an error; the worker cannot be stopped and is left running.
    /// </summary>
    public YamlConformanceChecks(IYamlConformanceEngine engine, TimeSpan? timeout = null) {
        _engine = engine ?? throw new ArgumentNullException(nameof(engine));
        _timeout = timeout;
    }

    /// <summary>Runs every applicable check on a suite case. Round-trips run only on valid cases.</summary>
    public IReadOnlyList<YamlCheckResult> CheckSuiteCase(YamlTestSuiteCase testCase) {
        if (testCase == null) throw new ArgumentNullException(nameof(testCase));
        var results = new List<YamlCheckResult> { CheckEvents(testCase), CheckErrorCase(testCase), CheckJson(testCase) };
        if (!testCase.IsError) {
            results.Add(CheckRoundTrip(testCase.Id, testCase.Input));
            results.Add(CheckRoundTripContent(testCase.Id, testCase.Input));
        }

        return results;
    }

    /// <summary>Runs the acceptance and round-trip checks on a corpus file's YAML.</summary>
    public IReadOnlyList<YamlCheckResult> CheckCorpusFile(YamlCorpusFile file) {
        if (file == null) throw new ArgumentNullException(nameof(file));
        return new[] { CheckAccepts(file.Path, file.Yaml), CheckRoundTrip(file.Path, file.Yaml), CheckRoundTripContent(file.Path, file.Yaml) };
    }

    /// <summary>Compares the events of a valid case with <c>test.event</c>.</summary>
    public YamlCheckResult CheckEvents(YamlTestSuiteCase testCase) {
        if (testCase == null) throw new ArgumentNullException(nameof(testCase));
        const YamlConformanceCheck check = YamlConformanceCheck.Events;
        if (testCase.IsError || !Has(YamlEngineFeatures.Events) || Has(YamlEngineFeatures.ContentEvents)) return Skip(testCase.Id, check);
        if (!TryRun(() => _engine.Parse(testCase.Input), out IReadOnlyList<YamlTestEvent>? events, out string? error)) {
            return new YamlCheckResult(testCase.Id, check, YamlCheckOutcome.Error, error);
        }

        IReadOnlyList<string> expected = YamlTestEventText.Normalize(testCase.ExpectedEvents);
        IReadOnlyList<string> actual = YamlTestEventText.Format(events!);
        int count = Math.Min(expected.Count, actual.Count);
        for (int i = 0; i < count; i++) {
            if (!string.Equals(expected[i], actual[i], StringComparison.Ordinal)) {
                return new YamlCheckResult(testCase.Id, check, YamlCheckOutcome.Fail, "line " + (i + 1) + ": expected '" + expected[i] + "' but got '" + actual[i] + "'");
            }
        }

        if (expected.Count != actual.Count) {
            return new YamlCheckResult(testCase.Id, check, YamlCheckOutcome.Fail, "expected " + expected.Count + " events but got " + actual.Count);
        }

        return new YamlCheckResult(testCase.Id, check, YamlCheckOutcome.Pass);
    }

    /// <summary>Checks that an invalid case is rejected, by the parser when available and otherwise by the loader.</summary>
    public YamlCheckResult CheckErrorCase(YamlTestSuiteCase testCase) {
        if (testCase == null) throw new ArgumentNullException(nameof(testCase));
        const YamlConformanceCheck check = YamlConformanceCheck.ErrorCase;
        if (!testCase.IsError || (!Has(YamlEngineFeatures.Events) && !Has(YamlEngineFeatures.Load))) return Skip(testCase.Id, check);
        bool completed = Has(YamlEngineFeatures.Events)
            ? TryRun(() => _engine.Parse(testCase.Input), out _, out string? error)
            : TryRun(() => _engine.Load(testCase.Input), out _, out error);
        if (error != null && error.StartsWith(TimeoutPrefix, StringComparison.Ordinal)) return new YamlCheckResult(testCase.Id, check, YamlCheckOutcome.Error, error);
        return completed
            ? new YamlCheckResult(testCase.Id, check, YamlCheckOutcome.Fail, "accepted invalid input")
            : new YamlCheckResult(testCase.Id, check, YamlCheckOutcome.Pass);
    }

    /// <summary>Compares the loaded documents of a valid case with <c>in.json</c>.</summary>
    public YamlCheckResult CheckJson(YamlTestSuiteCase testCase) {
        if (testCase == null) throw new ArgumentNullException(nameof(testCase));
        const YamlConformanceCheck check = YamlConformanceCheck.Json;
        if (testCase.IsError || testCase.Json == null || !Has(YamlEngineFeatures.Load)) return Skip(testCase.Id, check);
        if (!TryRun(() => _engine.Load(testCase.Input), out IReadOnlyList<YamlConformanceNode>? documents, out string? error)) {
            return new YamlCheckResult(testCase.Id, check, YamlCheckOutcome.Error, error);
        }

        string? difference = YamlJsonEquivalence.FindDifference(YamlJsonReader.ReadStream(testCase.Json), documents!);
        return difference == null
            ? new YamlCheckResult(testCase.Id, check, YamlCheckOutcome.Pass)
            : new YamlCheckResult(testCase.Id, check, YamlCheckOutcome.Fail, difference);
    }

    /// <summary>Checks that the engine parses (or, without events, loads) the input.</summary>
    public YamlCheckResult CheckAccepts(string subject, byte[] yaml) {
        const YamlConformanceCheck check = YamlConformanceCheck.Accepts;
        if (!Has(YamlEngineFeatures.Events) && !Has(YamlEngineFeatures.Load)) return Skip(subject, check);
        bool completed = Has(YamlEngineFeatures.Events)
            ? TryRun(() => _engine.Parse(yaml), out _, out string? error)
            : TryRun(() => _engine.Load(yaml), out _, out error);
        return completed ? new YamlCheckResult(subject, check, YamlCheckOutcome.Pass) : new YamlCheckResult(subject, check, YamlCheckOutcome.Error, error);
    }

    /// <summary>Checks that an unedited load and save returns the input bytes.</summary>
    public YamlCheckResult CheckRoundTrip(string subject, byte[] yaml) {
        const YamlConformanceCheck check = YamlConformanceCheck.RoundTrip;
        if (!Has(YamlEngineFeatures.RoundTrip)) return Skip(subject, check);
        if (!TryRun(() => _engine.RoundTrip(yaml), out byte[]? output, out string? error)) return new YamlCheckResult(subject, check, YamlCheckOutcome.Error, error);
        int difference = FirstDifference(yaml, output!);
        return difference < 0
            ? new YamlCheckResult(subject, check, YamlCheckOutcome.Pass)
            : new YamlCheckResult(subject, check, YamlCheckOutcome.Fail, "first difference at byte " + difference.ToString(CultureInfo.InvariantCulture) + " of " + yaml.Length.ToString(CultureInfo.InvariantCulture));
    }

    /// <summary>
    /// Checks that an unedited load and save keeps the content, by parsing input and output with the same engine and
    /// comparing values, tags and anchors while ignoring presentation.
    /// </summary>
    public YamlCheckResult CheckRoundTripContent(string subject, byte[] yaml) {
        const YamlConformanceCheck check = YamlConformanceCheck.RoundTripContent;
        if (!Has(YamlEngineFeatures.RoundTrip) || !Has(YamlEngineFeatures.Events)) return Skip(subject, check);
        if (!TryRun(() => _engine.Parse(yaml), out IReadOnlyList<YamlTestEvent>? original, out string? error)) {
            return new YamlCheckResult(subject, check, YamlCheckOutcome.Skipped, "input not parsed: " + error);
        }

        if (!TryRun(() => _engine.Parse(_engine.RoundTrip(yaml)), out IReadOnlyList<YamlTestEvent>? saved, out error)) {
            return new YamlCheckResult(subject, check, YamlCheckOutcome.Error, error);
        }

        int index = YamlTestEventText.FindContentDifference(original!, saved!);
        if (index < 0) return new YamlCheckResult(subject, check, YamlCheckOutcome.Pass);
        string expected = index < original!.Count ? original[index].ToString() : "<end>";
        string actual = index < saved!.Count ? saved[index].ToString() : "<end>";
        return new YamlCheckResult(subject, check, YamlCheckOutcome.Fail, "event " + index + ": expected '" + expected + "' but re-parsed '" + actual + "'");
    }

    private const string TimeoutPrefix = "timed out";

    private bool Has(YamlEngineFeatures feature) => (_engine.Features & feature) == feature;

    private static YamlCheckResult Skip(string subject, YamlConformanceCheck check) => new(subject, check, YamlCheckOutcome.Skipped);

    private bool TryRun<T>(Func<T> action, out T? result, out string? error) where T : class {
        result = null;
        error = null;
        try {
            if (_timeout == null) {
                result = action();
                return true;
            }

            Task<T> task = Task.Run(action);
            if (!task.Wait(_timeout.Value)) {
                error = TimeoutPrefix + " after " + _timeout.Value.TotalSeconds.ToString(CultureInfo.InvariantCulture) + " s";
                return false;
            }

            result = task.Result;
            return true;
        } catch (Exception exception) {
            Exception inner = exception is AggregateException aggregate && aggregate.InnerExceptions.Count == 1 ? aggregate.InnerExceptions[0] : exception;
            error = Describe(inner);
            return false;
        }
    }

    private static string Describe(Exception exception) {
        string message = exception.Message;
        int newline = message.IndexOf('\n');
        if (newline >= 0) message = message.Substring(0, newline).TrimEnd('\r');
        if (message.Length > 160) message = message.Substring(0, 160) + "...";
        return exception.GetType().Name + ": " + message;
    }

    private static int FirstDifference(byte[] expected, byte[] actual) {
        int count = Math.Min(expected.Length, actual.Length);
        for (int i = 0; i < count; i++) {
            if (expected[i] != actual[i]) return i;
        }

        return expected.Length == actual.Length ? -1 : count;
    }
}
