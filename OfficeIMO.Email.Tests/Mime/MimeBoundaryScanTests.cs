namespace OfficeIMO.Email.Tests;

public sealed class MimeBoundaryScanTests {
    [Theory]
    [InlineData(65535, "\r\n", true)]
    [InlineData(65536, "\n", true)]
    [InlineData(65537, "\r", true)]
    [InlineData(65535, "\r\n", false)]
    [InlineData(65536, "\n", false)]
    [InlineData(65537, "\r", false)]
    public void WindowCrossingsKeepCompleteDelimiterRulesAndFinalPartRecovery(
        int boundaryOffset, string newline, bool closed) {
        string first = "one" + newline + "x--edge" + newline + "--edgeX" + newline + "--edge-";
        string source = new string('p', boundaryOffset - newline.Length) + newline +
            "--edge \t" + newline + first + newline + "--edge" + newline + "two" +
            (closed ? newline + "--edge--\t" + newline + "epilogue" : string.Empty);
        byte[] bytes = Encoding.ASCII.GetBytes(source);
        var diagnostics = new List<EmailDiagnostic>();
        var state = new MimeParserState(new EmailReaderOptions(), diagnostics, CancellationToken.None);

        List<ArraySegment<byte>> parts = MimeParser.SplitMultipart(bytes, 0, bytes.Length, "edge", state, "message");

        Assert.Equal(new[] { first, "two" }, parts.Select(part =>
            Encoding.ASCII.GetString(part.Array!, part.Offset, part.Count)));
        Assert.Equal(closed ? Array.Empty<string>() : new[] { MimeParser.BoundaryNotClosedDiagnosticCode },
            diagnostics.Select(diagnostic => diagnostic.Code));
    }

    [Fact]
    public void NonzeroRangeKeepsEmptyPartsAndTruncatedMarkersWithoutReadingOutsideIt() {
        const string range = "--b\n\n--b\n--";
        byte[] bytes = Encoding.ASCII.GetBytes("prefix?" + range + "b--\n");
        var diagnostics = new List<EmailDiagnostic>();
        var state = new MimeParserState(new EmailReaderOptions(), diagnostics, CancellationToken.None);

        List<ArraySegment<byte>> parts = MimeParser.SplitMultipart(bytes, 7, range.Length, "b", state, "message");

        Assert.Equal(new[] { string.Empty, "--" }, parts.Select(part =>
            Encoding.ASCII.GetString(part.Array!, part.Offset, part.Count)));
        Assert.Equal(MimeParser.BoundaryNotClosedDiagnosticCode, Assert.Single(diagnostics).Code);
    }
}
