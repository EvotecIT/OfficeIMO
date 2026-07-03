using OfficeIMO.Markdown;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

internal static class MarkdownSpecSyntaxAssert {
    public static void AssertSyntaxAssertions(MarkdownSyntaxNode root, IReadOnlyList<MarkdownSpecSyntaxAssertionFixture>? assertions) {
        if (assertions == null || assertions.Count == 0) {
            return;
        }

        foreach (var assertion in assertions) {
            var path = root.FindNodePathAtPosition(assertion.Line, assertion.Column);
            Assert.True(
                path.Count > 0,
                $"Expected a syntax path at L{assertion.Line}:C{assertion.Column} for '{assertion.Description ?? "spec assertion"}'.");

            if (assertion.ExpectedPathKinds is { Length: > 0 }) {
                Assert.Equal(assertion.ExpectedPathKinds, path.Select(node => node.Kind.ToString()).ToArray());
            }

            var deepest = path[path.Count - 1];

            if (!string.IsNullOrWhiteSpace(assertion.ExpectedDeepestKind)) {
                Assert.Equal(assertion.ExpectedDeepestKind, deepest.Kind.ToString());
            }

            if (assertion.ExpectedDeepestSpan != null) {
                Assert.Equal(assertion.ExpectedDeepestSpan.ToSourceSpan(), deepest.SourceSpan);
            }

            if (assertion.ExpectedDeepestLiteral != null) {
                Assert.Equal(assertion.ExpectedDeepestLiteral, deepest.Literal);
            }

            if (assertion.ExpectedNearestBlockKind == null
                && assertion.ExpectedNearestBlockSpan == null
                && assertion.ExpectedNearestBlockLiteral == null) {
                continue;
            }

            var nearestBlock = root.FindNearestBlockAtPosition(assertion.Line, assertion.Column);
            Assert.NotNull(nearestBlock);

            if (!string.IsNullOrWhiteSpace(assertion.ExpectedNearestBlockKind)) {
                Assert.Equal(assertion.ExpectedNearestBlockKind, nearestBlock!.Kind.ToString());
            }

            if (assertion.ExpectedNearestBlockSpan != null) {
                Assert.Equal(assertion.ExpectedNearestBlockSpan.ToSourceSpan(), nearestBlock!.SourceSpan);
            }

            if (assertion.ExpectedNearestBlockLiteral != null) {
                Assert.Equal(assertion.ExpectedNearestBlockLiteral, nearestBlock!.Literal);
            }
        }
    }
}

public sealed class MarkdownSpecSyntaxAssertionFixture {
    public string? Description { get; set; }
    public int Line { get; set; }
    public int Column { get; set; }
    public string[]? ExpectedPathKinds { get; set; }
    public string? ExpectedDeepestKind { get; set; }
    public string? ExpectedDeepestLiteral { get; set; }
    public MarkdownSpecSourceSpanFixture? ExpectedDeepestSpan { get; set; }
    public string? ExpectedNearestBlockKind { get; set; }
    public string? ExpectedNearestBlockLiteral { get; set; }
    public MarkdownSpecSourceSpanFixture? ExpectedNearestBlockSpan { get; set; }
}

public sealed class MarkdownSpecSourceSpanFixture {
    public int StartLine { get; set; }
    public int StartColumn { get; set; }
    public int EndLine { get; set; }
    public int EndColumn { get; set; }

    public MarkdownSourceSpan ToSourceSpan() => new(StartLine, StartColumn, EndLine, EndColumn);
}
