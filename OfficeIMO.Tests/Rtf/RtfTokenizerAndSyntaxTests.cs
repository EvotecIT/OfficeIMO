using OfficeIMO.Rtf.Diagnostics;
using OfficeIMO.Rtf.Syntax;
using Xunit;

namespace OfficeIMO.Tests.Rtf;

public class RtfTokenizerAndSyntaxTests {
    [Fact]
    public void Tokenizer_Recognizes_ControlWords_Text_Hex_And_Binary() {
        RtfTokenizeResult result = RtfTokenizer.Tokenize(@"{\rtf1\ansi Hello \'80{\bin3 abc}}");

        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == RtfDiagnosticSeverity.Error);
        Assert.Contains(result.Tokens, token => token.Kind == RtfTokenKind.ControlWord && token.ControlName == "rtf" && token.Parameter == 1);
        Assert.Contains(result.Tokens, token => token.Kind == RtfTokenKind.ControlSymbol && token.ControlSymbol == '\'' && token.Parameter == 0x80);
        RtfToken binary = Assert.Single(result.Tokens, token => token.Kind == RtfTokenKind.Binary);
        Assert.Equal(new byte[] { 97, 98, 99 }, binary.BinaryData);
    }

    [Fact]
    public void SyntaxParser_Builds_Nested_Groups_And_Diagnostics_Unclosed_Input() {
        RtfSyntaxTree tree = RtfSyntaxTree.Parse(@"{\rtf1{\b Bold}");

        Assert.Equal("rtf", tree.Root.Destination);
        Assert.Contains(tree.Root.Children, node => node is RtfGroup);
        Assert.Contains(tree.Diagnostics, diagnostic => diagnostic.Code == "RTF013");
    }
}
