using OfficeIMO.Yaml.Tests.Conformance;
using Xunit;

namespace OfficeIMO.Yaml.Tests;

// The pinned yaml-test-suite is the reader's main proof, in both modes: every valid case must produce test.event exactly,
// and every error case must be rejected in strict mode. The default mode also accepts the error cases in
// LenientAccepted, and no others. Known gaps belong in Exceptions with a reason; the list only shrinks.
public class YamlReaderConformanceTests {
    /// <summary>Cases OfficeIMO.Yaml is known to get wrong, with the reason. A listed case that starts passing fails the test.</summary>
    private static readonly IReadOnlyDictionary<string, string> Exceptions = new Dictionary<string, string>(StringComparer.Ordinal);

    /// <summary>
    /// Error cases the default (non-strict) mode accepts, with the reason. libyaml and SharpYaml accept each of
    /// them and real files rely on the same forms. A listed case the default mode rejects, or an unlisted error case it
    /// accepts, fails the test.
    /// </summary>
    private static readonly IReadOnlyDictionary<string, string> LenientAccepted = new Dictionary<string, string>(StringComparer.Ordinal) {
        ["9C9N"] = "flow sequence lines at the parent key's indentation",
        ["QB6E"] = "double-quoted continuation lines at the parent key's indentation",
        ["VJP3/00"] = "flow mapping lines, including the closing brace, at the parent key's indentation"
    };

    private static readonly Lazy<YamlTestSuite> Suite = new(() => YamlTestSuite.Load());

    public static TheoryData<string> CaseIds() {
        var data = new TheoryData<string>();
        foreach (YamlTestSuiteCase testCase in Suite.Value.Cases) data.Add(testCase.Id);
        return data;
    }

    public static TheoryData<string, bool> CaseIdsInBothModes() {
        var data = new TheoryData<string, bool>();
        foreach (YamlTestSuiteCase testCase in Suite.Value.Cases) {
            data.Add(testCase.Id, true);
            data.Add(testCase.Id, false);
        }
        return data;
    }

    [Theory]
    [MemberData(nameof(CaseIdsInBothModes))]
    public void ReaderMatchesTheSuite(string id, bool strict) {
        YamlTestSuiteCase testCase = Suite.Value.Cases.Single(c => c.Id == id);
        var engine = new OfficeIMOYamlConformanceEngine(strict);
        if (!strict && LenientAccepted.TryGetValue(id, out string? accepted)) {
            Assert.True(testCase.IsError, id + " is valid; only error cases belong in the lenient list (" + accepted + ").");
            IReadOnlyList<string> actual = YamlTestEventText.Format(engine.Parse(testCase.Input));
            // The events before the point the specification rejects must still match.
            IReadOnlyList<string> expected = YamlTestEventText.Normalize(testCase.ExpectedEvents);
            Assert.True(actual.Count > expected.Count, id + " stopped after " + actual.Count + " events.");
            for (int i = 0; i < expected.Count; i++) Assert.Equal(expected[i], actual[i]);
            return;
        }

        var checks = new YamlConformanceChecks(engine);
        YamlCheckResult result = testCase.IsError ? checks.CheckErrorCase(testCase) : checks.CheckEvents(testCase);
        if (Exceptions.TryGetValue(id, out string? reason)) {
            Assert.True(result.Outcome != YamlCheckOutcome.Pass, id + " now passes; remove it from the exception list (" + reason + ").");
            return;
        }

        Assert.True(result.Outcome == YamlCheckOutcome.Pass, result.ToString());
    }

    // Positions are the lossless contract the document model rebuilds from: event spans never overlap, and the text
    // between them is only whitespace, line breaks, comments and the indicators that produce no event.
    [Theory]
    [MemberData(nameof(CaseIds))]
    public void EventSpansLeaveOnlyTriviaBetweenThem(string id) {
        YamlTestSuiteCase testCase = Suite.Value.Cases.Single(c => c.Id == id);
        if (testCase.IsError) return;
        string text = YamlTextTestHelpers.DecodeSuiteInput(testCase.Input);
        using var reader = new YamlReader(text, new YamlReaderOptions { IncludeComments = true });
        long covered = 0;
        while (reader.Read()) {
            long start = reader.Start.Offset;
            long end = reader.End.Offset;
            Assert.True(end >= start, id + ": " + reader.TokenType + " ends before it starts.");
            // Comments inside a directive line or a block scalar header lie within the enclosing event's span.
            if (reader.TokenType == YamlTokenType.Comment && start < covered) continue;
            Assert.True(start >= covered, id + ": " + reader.TokenType + " at " + reader.Start + " overlaps the previous event.");
            string gap = text.Substring((int)covered, (int)(start - covered));
            Assert.True(IsTrivia(gap), id + ": unexpected text '" + gap + "' before " + reader.TokenType + " at " + reader.Start + ".");
            covered = end;
        }

        Assert.True(IsTrivia(text.Substring((int)covered)), id + ": unexpected text after the last event.");
    }

    // The suite cannot say how accepted indentation is reported: once per outermost flow collection and once per quoted
    // scalar, at the first shallow line, by the reader and the model alike; a clean file reports nothing and allocates nothing.
    private const string Shallow = "a: [\n  1,\n  2\n]\nb: {\nx: 1,\ny: [2,\n3]\n}\nc: 'one\ntwo\nthree'\nd: [\n  ok\n  ]\n";

    [Fact]
    public void AcceptedShallowIndentationIsReportedOncePerConstruct() {
        using var reader = new YamlReader(Shallow);
        Assert.Same(Array.Empty<OfficeDiagnostic>(), reader.Diagnostics);
        while (reader.Read()) { }
        Assert.Equal(new[] { "4:1", "6:1", "11:1" }, reader.Diagnostics.Select(d => d.Location));
        Assert.All(reader.Diagnostics, d => {
            Assert.Equal(OfficeDiagnosticCodes.YamlLenientIndentation, d.Code);
            Assert.Equal(OfficeDiagnosticSeverity.Warning, d.Severity);
        });

        YamlDocument document = YamlDocument.Parse(Shallow);
        Assert.Equal(new[] { "4:1", "6:1", "11:1" }, document.Diagnostics.Select(d => d.Location));
        Assert.Equal("one two three", document.Select("c")!.AsString());
        Assert.Equal(YamlErrorCodes.InvalidIndentation, Assert.Throws<YamlException>(() => YamlDocument.Parse(Shallow, new YamlLoadOptions { Strict = true })).Code);

        using var clean = new YamlReader("a: [1,\n  2]\nb: 'x\n  y'\n");
        while (clean.Read()) { }
        Assert.Same(Array.Empty<OfficeDiagnostic>(), clean.Diagnostics);
    }

    // Saving copies accepted shallow text it does not own and never writes it: an edited quoted scalar and the
    // separator of a new flow item take the valid indentation.
    [Fact]
    public void EditsNearShallowIndentationWriteValidText() {
        YamlDocument document = YamlDocument.Parse(Shallow);
        document.Set("c", "one two four");
        document.Add("a", 3);
        Assert.Equal("a: [\n  1,\n  2,\n  3\n]\nb: {\nx: 1,\ny: [2,\n3]\n}\nc: 'one two four'\nd: [\n  ok\n  ]\n", document.ToString());
    }

    // The suite has no comment at the start of the line after a block scalar, whose span ends at that line start.
    [Theory]
    [InlineData("a: |\n  x\n# c\nb: 1\n")]
    [InlineData("- >-\n  x\n\n# c\n- y\n")]
    public void CommentAtTheStartOfTheLineAfterABlockScalarIsTrivia(string yaml) {
        using var reader = new YamlReader(yaml);
        int scalars = 0;
        while (reader.Read()) {
            if (reader.TokenType == YamlTokenType.Scalar) scalars++;
        }
        Assert.Equal(yaml.StartsWith("a", StringComparison.Ordinal) ? 4 : 2, scalars);
    }

    private static bool IsTrivia(string gap) {
        foreach (char ch in gap) {
            if (ch is not (' ' or '\t' or '\r' or '\n' or '-' or '?' or ':' or ',' or '.' or '﻿')) return false;
        }

        return true;
    }
}
