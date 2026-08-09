using OfficeIMO.Rtf;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public sealed class RtfFieldCodeSyntaxTests {
    [Fact]
    public void HyperlinkProjectionDecodesEscapedQuotesWithoutLosingFollowingWords() {
        const string instruction = "HYPERLINK \"https://example.test/\" \\o \"He said \\\"Hi\\\" today\"";

        var field = new RtfField(instruction);

        Assert.True(field.FieldCode.IsValid);
        Assert.Equal("HYPERLINK", field.FieldCode.Keyword);
        Assert.Equal(new Uri("https://example.test/"), field.Hyperlink);
        Assert.Equal("He said \"Hi\" today", field.HyperlinkField!.ScreenTip);
        Assert.Equal(instruction, field.HyperlinkField.ToInstruction());
        Assert.Equal(instruction, string.Concat(field.FieldCode.Tokens.Select(static token => token.Text)));
    }

    [Fact]
    public void UnknownFieldSwitchesRemainLosslessTokens() {
        const string instruction = "MERGEFIELD Customer \\* MERGEFORMAT \\x \"future value\"";

        RtfFieldCodeSyntax syntax = RtfFieldCodeSyntax.Parse(instruction);

        Assert.True(syntax.IsValid);
        Assert.Equal("MERGEFIELD", syntax.Keyword);
        Assert.Contains(syntax.Tokens, static token => token.Kind == RtfFieldCodeTokenKind.Switch && token.Value == "x");
        Assert.Equal(instruction, string.Concat(syntax.Tokens.Select(static token => token.Text)));
    }

    [Fact]
    public void UnterminatedQuotedArgumentIsPreservedAndInvalid() {
        const string instruction = "HYPERLINK \"https://example.test/";

        RtfFieldCodeSyntax syntax = RtfFieldCodeSyntax.Parse(instruction);

        Assert.False(syntax.IsValid);
        Assert.Equal(instruction, string.Concat(syntax.Tokens.Select(static token => token.Text)));
    }

    [Fact]
    public void HyperlinkFormattingSwitchArgumentIsNotMistakenForTarget() {
        const string instruction = "HYPERLINK \\l \"Bookmark\" \\* MERGEFORMAT";

        var field = new RtfField(instruction);

        Assert.NotNull(field.HyperlinkField);
        Assert.Equal("Bookmark", field.HyperlinkField!.SubAddress);
        Assert.Null(field.HyperlinkField.Target);
    }
}