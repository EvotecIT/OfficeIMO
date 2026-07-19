using OfficeIMO.Drawing;
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
    public void Interpreter_SkipsOperationsWithUnrepresentableNumbersWithoutShiftingOperands() {
        const string content =
            "1e309 40 20 20 re " +
            "5 1e309 40 20 20 re " +
            "5 40 1e309 20 20 re " +
            "5 40 20 1e309 20 re " +
            "5 40 20 20 1e309 re " +
            "5 6 m";
        var operations = new List<PdfContentOperation>();

        PdfContentStreamInterpreter.Interpret(content, 10, operations.Add);

        PdfContentOperation operation = Assert.Single(operations);
        Assert.Equal("m", operation.Name);
        Assert.Equal(new[] { 5D, 6D }, operation.Operands.Cast<double>());
    }

    [Fact]
    public void Interpreter_SkipsInvalidCompoundOperandsWithoutLosingSynchronization() {
        const string content =
            "[1 1e309 2] TJ " +
            "[1e309] Q " +
            "<< /X 1e309 >> n " +
            "/Span << /MCID 1e309 >> BDC EMC " +
            "BI /W 1e309 /H 1 /BPC 8 /CS /G ID A EI " +
            "5 6 m";
        var operations = new List<PdfContentOperation>();

        PdfContentStreamInterpreter.Interpret(content, 10, operations.Add);

        Assert.Equal(new[] { "Q", "n", "BDC", "EMC", "m" }, operations.Select(operation => operation.Name));
        Assert.All(operations.Take(4), operation => Assert.Empty(operation.Operands));
        Assert.True(operations[2].HasInvalidOperands);
        Assert.Equal(new[] { 5D, 6D }, operations[4].Operands.Cast<double>());
    }

    [Fact]
    public void Interpreter_PreservesValidInlineImageAfterInvalidSurplusOperand() {
        const string content = "1e309 BI /W 1 /H 1 /BPC 8 /CS /G ID A EI q Q";
        var operations = new List<PdfContentOperation>();

        PdfContentStreamInterpreter.Interpret(content, 10, operations.Add);

        Assert.Equal(new[] { "BI", "q", "Q" }, operations.Select(operation => operation.Name));
        PdfContentInlineImage inlineImage = Assert.IsType<PdfContentInlineImage>(operations[0].InlineImage);
        Assert.True(operations[0].HasInvalidOperands);
        Assert.Equal(new byte[] { (byte)'A' }, inlineImage.Data);
        Assert.Equal(
            1,
            PdfPageXObjectInvocationParser.Parse(content, Matrix2D.Identity, 200D).Count);
    }

    [Fact]
    public void InterpreterUntil_VisitsOperandlessRecoveryOperatorAfterInvalidOperand() {
        var operations = new List<PdfContentOperation>();

        bool completed = PdfContentStreamInterpreter.InterpretUntil(
            "1e309 Q 5 6 m",
            10,
            operation => {
                operations.Add(operation);
                return !string.Equals(operation.Name, "Q", StringComparison.Ordinal);
            });

        Assert.False(completed);
        PdfContentOperation operation = Assert.Single(operations);
        Assert.Equal("Q", operation.Name);
        Assert.Empty(operation.Operands);
    }

    [Fact]
    public void TextParser_DoesNotCombineOperandsAcrossSkippedMalformedOperations() {
        const string content = "BT /F2 Tf 1e309 cm 18 Tf 1 0 0 1 0 0 Tm (A) Tj ET";

        List<PdfTextSpan> spans = TextContentParser.Parse(
            content,
            (_, bytes) => System.Text.Encoding.ASCII.GetString(bytes),
            (_, bytes) => bytes.Length * 500D);

        PdfTextSpan span = Assert.Single(spans);
        Assert.Equal("F1", span.FontResource);
        Assert.Equal(12D, span.FontSize);
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
    public void Interpreter_StopsNestedOperandsBeforeRecursiveDescentExceedsBudget() {
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfContentStreamInterpreter.Interpret(
                "[[ [1] ]] TJ",
                10,
                _ => { },
                maxNestingDepth: 2));

        Assert.Equal(PdfReadLimitKind.ContentNestingDepth, exception.Kind);
        Assert.Equal(2, exception.Limit);
        Assert.Equal(3, exception.Actual);
    }

    [Fact]
    public void Interpreter_StopsFlatOperandAmplificationBeforeAnOperator() {
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfContentStreamInterpreter.Interpret(
                "1 2 3 4 5 6 7 8 9 10 cm",
                10,
                _ => { },
                maxOperands: 4));

        Assert.Equal(PdfReadLimitKind.ContentOperands, exception.Kind);
        Assert.Equal(4, exception.Limit);
        Assert.Equal(5, exception.Actual);
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

    [Fact]
    public void OperatorScanner_PreservesInlineImageFramingOperators() {
        const string content = "q BI /W 1 /H 1 /BPC 8 /CS /G ID A EI Q";
        var operators = new List<string>();
        bool truncated = false;

        PdfContentOperatorScanner.AppendOperators(content, operators, 10, ref truncated);

        Assert.Equal(new[] { "q", "BI", "ID", "EI", "Q" }, operators);
        Assert.False(truncated);
    }
}
