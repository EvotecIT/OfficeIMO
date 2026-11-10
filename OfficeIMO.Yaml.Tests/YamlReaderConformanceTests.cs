using OfficeIMO.Yaml.Tests.Conformance;
using Xunit;

namespace OfficeIMO.Yaml.Tests;

// The pinned yaml-test-suite is the reader's main proof: every valid case must produce test.event exactly and every
// error case must be rejected. Known gaps belong in Exceptions with a reason; the list only shrinks.
public class YamlReaderConformanceTests {
    /// <summary>Cases OfficeIMO.Yaml is known to get wrong, with the reason. A listed case that starts passing fails the test.</summary>
    private static readonly IReadOnlyDictionary<string, string> Exceptions = new Dictionary<string, string>(StringComparer.Ordinal);

    private static readonly Lazy<YamlTestSuite> Suite = new(() => YamlTestSuite.Load());

    public static TheoryData<string> CaseIds() {
        var data = new TheoryData<string>();
        foreach (YamlTestSuiteCase testCase in Suite.Value.Cases) data.Add(testCase.Id);
        return data;
    }

    [Theory]
    [MemberData(nameof(CaseIds))]
    public void ReaderMatchesTheSuite(string id) {
        YamlTestSuiteCase testCase = Suite.Value.Cases.Single(c => c.Id == id);
        var checks = new YamlConformanceChecks(new OfficeIMOYamlConformanceEngine());
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

    private static bool IsTrivia(string gap) {
        foreach (char ch in gap) {
            if (ch is not (' ' or '\t' or '\r' or '\n' or '-' or '?' or ':' or ',' or '.' or '﻿')) return false;
        }

        return true;
    }
}
