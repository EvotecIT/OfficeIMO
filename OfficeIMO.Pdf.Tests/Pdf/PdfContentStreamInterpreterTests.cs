using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfContentStreamInterpreterTests {
    [Fact]
    public void Interpreter_EmitsTypedOperandsForAllVisitors() {
        const string content =
            "% comment\n" +
            "/Span << /ActualText (logical) /OC /LayerOne >> BDC " +
            "BT /F1 12 Tf [(A) -20 <42>] TJ ET EMC";
        var operations = new List<PdfContentOperation>();

        PdfContentStreamInterpreter.Interpret(content, 20, operations.Add);

        Assert.Equal(new[] { "BDC", "BT", "Tf", "TJ", "ET", "EMC" }, operations.Select(operation => operation.Name));
        PdfContentDictionary dictionary = Assert.IsType<PdfContentDictionary>(operations[0].Operands[1]);
        Assert.Equal("logical", PdfTextString.Decode(Assert.IsType<byte[]>(dictionary.Items["ActualText"])));
        Assert.Equal("LayerOne", Assert.IsType<string>(dictionary.Items["OC"]));
        List<object> textArray = Assert.IsType<List<object>>(operations[3].Operands[0]);
        Assert.Equal(3, textArray.Count);
        Assert.Equal("A", PdfTextString.Decode(Assert.IsType<byte[]>(textArray[0])));
        Assert.Equal(-20D, Assert.IsType<double>(textArray[1]));
        Assert.Equal("B", PdfTextString.Decode(Assert.IsType<byte[]>(textArray[2])));
    }

    [Fact]
    public void Interpreter_UsesRawInlineImageLengthBeforeTerminatorHeuristics() {
        const string content = "q BI /W 4 /H 1 /BPC 8 /CS /G ID A EI EI Q";
        var operations = new List<PdfContentOperation>();

        PdfContentStreamInterpreter.Interpret(content, 10, operations.Add);

        Assert.Equal(new[] { "q", "BI", "Q" }, operations.Select(operation => operation.Name));
        PdfContentInlineImage inlineImage = Assert.IsType<PdfContentInlineImage>(operations[1].InlineImage);
        Assert.Equal(new byte[] { (byte)'A', (byte)' ', (byte)'E', (byte)'I' }, inlineImage.Data);
        Assert.Equal(4D, Assert.IsType<PdfNumber>(inlineImage.Dictionary.Items["Width"]).Value);
        Assert.Equal("DeviceGray", Assert.IsType<PdfName>(inlineImage.Dictionary.Items["ColorSpace"]).Name);
    }

    [Fact]
    public void Interpreter_AppliesOneSharedOperationBudget() {
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfContentStreamInterpreter.Interpret("q Q q", 2, _ => { }));

        Assert.Equal(PdfReadLimitKind.ContentOperations, exception.Kind);
        Assert.Equal(2, exception.Limit);
        Assert.Equal(3, exception.Actual);
    }

    [Fact]
    public void Interpreter_PreservesQuoteCharactersInNamesAndRecognizesQuoteOperators() {
        const string content = "/Owner's MP /A\"B MP (x) ' 1 2 (y) \"";
        var operations = new List<PdfContentOperation>();

        PdfContentStreamInterpreter.Interpret(content, 10, operations.Add);

        Assert.Equal(new[] { "MP", "MP", "'", "\"" }, operations.Select(operation => operation.Name));
        Assert.Equal("Owner's", Assert.IsType<string>(operations[0].Operands[0]));
        Assert.Equal("A\"B", Assert.IsType<string>(operations[1].Operands[0]));
    }
}
