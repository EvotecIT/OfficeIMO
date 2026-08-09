using OfficeIMO.Spreadsheet;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed class SpreadsheetFormulaSyntaxTests {
    [Theory]
    [InlineData("LOG10(100)", "of:=LOG10(100)")]
    [InlineData("=LOG10(A1)", "of:=LOG10([.A1])")]
    [InlineData("=SUM(A1,B1)", "of:=SUM([.A1];[.B1])")]
    [InlineData("={1,2;3,4}", "of:={1;2|3;4}")]
    [InlineData("=IF(A1=\"B2\",1,0)", "of:=IF([.A1]=\"B2\";1;0)")]
    [InlineData("=SUM((A1,B1))", "of:=SUM(([.A1]~[.B1]))")]
    [InlineData("=SUM(A1 B1)", "of:=SUM([.A1]![.B1])")]
    [InlineData("=MyRange A1", "of:=MyRange![.A1]")]
    [InlineData("=A1 OtherRange", "of:=[.A1]!OtherRange")]
    [InlineData("=FirstRange SecondRange", "of:=FirstRange!SecondRange")]
    [InlineData("=SUM($1:$2)", "of:=SUM([.$1:.$2])")]
    public void ExcelFormulaTranslationUsesStructuralContext(string excel, string expected) {
        SpreadsheetFormulaTranslationResult result = SpreadsheetFormulaSyntaxTree
            .Parse(excel, SpreadsheetFormulaDialect.ExcelA1)
            .TranslateTo(SpreadsheetFormulaDialect.OpenFormula);

        Assert.True(result.IsSuccessful, string.Join("; ", result.Diagnostics.Select(diagnostic => diagnostic.Message)));
        Assert.Equal(expected, result.Formula);
    }

    [Fact]
    public void WhitespaceBeforeAFunctionCallRemainsTriviaInsteadOfBecomingAnIntersection() {
        SpreadsheetFormulaTranslationResult result = SpreadsheetFormulaSyntaxTree
            .Parse("=A1+SUM (B1)", SpreadsheetFormulaDialect.ExcelA1)
            .TranslateTo(SpreadsheetFormulaDialect.OpenFormula);

        Assert.True(result.IsSuccessful, string.Join("; ", result.Diagnostics.Select(diagnostic => diagnostic.Message)));
        Assert.Equal("of:=[.A1]+SUM ([.B1])", result.Formula);
    }

    [Theory]
    [InlineData("of:=LOG10(100)", "=LOG10(100)")]
    [InlineData("of:=SUM([.A1];[.B1])", "=SUM(A1,B1)")]
    [InlineData("of:={1;2|3;4}", "={1,2;3,4}")]
    [InlineData("of:=SUM(([.A1]~[.B1]))", "=SUM((A1,B1))")]
    [InlineData("of:=SUM([.A1]![.B1])", "=SUM(A1 B1)")]
    [InlineData("of:=SUM([.$1:.$2])", "=SUM($1:$2)")]
    public void OpenFormulaTranslationUsesStructuralContext(string openFormula, string expected) {
        SpreadsheetFormulaTranslationResult result = SpreadsheetFormulaSyntaxTree
            .Parse(openFormula, SpreadsheetFormulaDialect.OpenFormula)
            .TranslateTo(SpreadsheetFormulaDialect.ExcelA1);

        Assert.True(result.IsSuccessful, string.Join("; ", result.Diagnostics.Select(diagnostic => diagnostic.Message)));
        Assert.Equal(expected, result.Formula);
    }

    [Fact]
    public void StructuredReferenceFailsClosedInsteadOfProducingInvalidOpenFormula() {
        SpreadsheetFormulaTranslationResult result = SpreadsheetFormulaSyntaxTree
            .Parse("=SUM(Table1[Amount])", SpreadsheetFormulaDialect.ExcelA1)
            .TranslateTo(SpreadsheetFormulaDialect.OpenFormula);

        Assert.False(result.IsSuccessful);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "FORMULA_TRANSLATION_UNSUPPORTED");
    }

    [Fact]
    public void ExcessiveFormulaNestingFailsClosedWithoutRecursiveOverflow() {
        string formula = "=" + new string('(', 5_000) + "A1" + new string(')', 5_000);

        SpreadsheetFormulaTranslationResult result = SpreadsheetFormulaSyntaxTree
            .Parse(formula, SpreadsheetFormulaDialect.ExcelA1)
            .TranslateTo(SpreadsheetFormulaDialect.OpenFormula);

        Assert.False(result.IsSuccessful);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "FORMULA_NESTING_LIMIT");
    }

    [Fact]
    public void ExcelThreeDimensionalReferenceFailsClosed() {
        SpreadsheetFormulaTranslationResult result = SpreadsheetFormulaSyntaxTree
            .Parse("=SUM(Sheet1:Sheet3!A1)", SpreadsheetFormulaDialect.ExcelA1)
            .TranslateTo(SpreadsheetFormulaDialect.OpenFormula);

        Assert.False(result.IsSuccessful);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "FORMULA_UNSUPPORTED_RANGE_OPERATOR");
    }

    [Theory]
    [InlineData("=#REF!A1")]
    [InlineData("=#REF!A1:C3")]
    public void ExcelDeletedReferencesWithAttachedAddressesFailClosed(string formula) {
        SpreadsheetFormulaTranslationResult result = SpreadsheetFormulaSyntaxTree
            .Parse(formula, SpreadsheetFormulaDialect.ExcelA1)
            .TranslateTo(SpreadsheetFormulaDialect.OpenFormula);

        Assert.False(result.IsSuccessful);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "FORMULA_DELETED_REFERENCE");
    }

    [Theory]
    [InlineData("=#N/A+A1", "of:=#N/A+[.A1]")]
    [InlineData("=#GETTING_DATA+A1", "of:=#GETTING_DATA+[.A1]")]
    public void ExcelErrorLiteralsStopBeforeFollowingOperators(string formula, string expected) {
        SpreadsheetFormulaTranslationResult result = SpreadsheetFormulaSyntaxTree
            .Parse(formula, SpreadsheetFormulaDialect.ExcelA1)
            .TranslateTo(SpreadsheetFormulaDialect.OpenFormula);

        Assert.True(result.IsSuccessful, string.Join("; ", result.Diagnostics.Select(diagnostic => diagnostic.Message)));
        Assert.Equal(expected, result.Formula);
    }

    [Fact]
    public void UnknownExcelErrorLiteralFailsClosed() {
        SpreadsheetFormulaTranslationResult result = SpreadsheetFormulaSyntaxTree
            .Parse("=#BOGUS+A1", SpreadsheetFormulaDialect.ExcelA1)
            .TranslateTo(SpreadsheetFormulaDialect.OpenFormula);

        Assert.False(result.IsSuccessful);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "FORMULA_UNKNOWN_ERROR_LITERAL");
    }

    [Theory]
    [InlineData("of:=[.A1048577]")]
    [InlineData("of:=[.XFE1]")]
    public void OpenFormulaReferencesOutsideExcelBoundsFailClosed(string formula) {
        SpreadsheetFormulaTranslationResult result = SpreadsheetFormulaSyntaxTree
            .Parse(formula, SpreadsheetFormulaDialect.OpenFormula)
            .TranslateTo(SpreadsheetFormulaDialect.ExcelA1);

        Assert.False(result.IsSuccessful);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == "FORMULA_TRANSLATION_REFERENCE_BOUNDS");
    }

    [Fact]
    public void OpenDocumentAddressParsesQuotedColonAndDerivesTypedBaseCell() {
        SpreadsheetRangeReference reference = SpreadsheetRangeReference.Parse(
            "$'A:B'.$C$1:.$C$3",
            SpreadsheetAddressDialect.OpenDocument);

        Assert.Equal("A:B", reference.Start.SheetName);
        Assert.Equal("$'A:B'.$C$1", reference.FormatBaseCell(SpreadsheetAddressDialect.OpenDocument));
        Assert.Equal("'A:B'!$C$1:$C$3", reference.Format(SpreadsheetAddressDialect.ExcelA1));
    }

    [Fact]
    public void ExcelQualifiedRangeInheritsAndRepeatsItsSheetInOpenDocumentSyntax() {
        SpreadsheetRangeReference reference = SpreadsheetRangeReference.Parse(
            "'Other'!A1:B2",
            SpreadsheetAddressDialect.ExcelA1);

        Assert.Equal("Other", reference.Start.SheetName);
        Assert.Equal("Other", reference.End!.SheetName);
        Assert.Equal("$'Other'.A1:$'Other'.B2", reference.Format(SpreadsheetAddressDialect.OpenDocument));
    }

    [Theory]
    [InlineData("Στοιχεία!A1", "Στοιχεία")]
    [InlineData("Données2!B3", "Données2")]
    [InlineData("Café!C4", "Café")]
    public void ExcelUnquotedSheetQualifierAcceptsUnicodeLettersMarksAndDigits(string address, string expectedSheet) {
        SpreadsheetRangeReference reference = SpreadsheetRangeReference.Parse(
            address,
            SpreadsheetAddressDialect.ExcelA1);

        Assert.Equal(expectedSheet, reference.Start.SheetName);
    }

    [Fact]
    public void UnboundedA1AcceptsCoordinatesOutsideExcelsGridWithoutChangingExcelRules() {
        Assert.True(SpreadsheetRangeReference.TryParse(
            "XFE1048577:XFF1048578",
            SpreadsheetAddressDialect.UnboundedA1,
            out SpreadsheetRangeReference? unbounded));
        Assert.Equal(16385, unbounded!.Start.Column);
        Assert.Equal(1048578L, unbounded.End!.Row);
        Assert.False(SpreadsheetRangeReference.TryParse(
            "XFE1048577:XFF1048578",
            SpreadsheetAddressDialect.ExcelA1,
            out _));
    }

    [Fact]
    public void ReferenceSequenceKeepsSeparatorsInsideQuotedSheetNames() {
        SpreadsheetReferenceSequence sequence = SpreadsheetReferenceSequence.Parse(
            "'Ops, Europe'!$A$1:$B$2 'Owner''s Data'!C3,D4",
            SpreadsheetAddressDialect.ExcelA1);

        Assert.Equal(3, sequence.References.Count);
        Assert.Equal("Ops, Europe", sequence.References[0].Start.SheetName);
        Assert.Equal("Owner's Data", sequence.References[1].Start.SheetName);
        Assert.Equal("D4", sequence.References[2].Format(SpreadsheetAddressDialect.ExcelA1));
    }

    [Fact]
    public void ReferenceSequenceRejectsUnterminatedQuotedSheetName() {
        Assert.False(SpreadsheetReferenceSequence.TryParse(
            "'Unclosed sheet!A1 B2", SpreadsheetAddressDialect.ExcelA1, out _));
    }
}

public sealed class SpreadsheetNumberFormatSyntaxTests {
    [Theory]
    [InlineData("0.00%", true, 2)]
    [InlineData("0.00\"%\"", false, 2)]
    [InlineData("#,##0.000", false, 3)]
    public void Parser_Distinguishes_Operators_From_Display_Literals(string format, bool percentage, int decimals) {
        SpreadsheetNumberFormatSyntax syntax = SpreadsheetNumberFormatSyntax.Parse(format);

        Assert.True(syntax.IsValid);
        Assert.Equal(percentage, syntax.IsPercentage);
        Assert.Equal(decimals, syntax.DecimalPlaces);
    }

    [Fact]
    public void Parser_Preserves_Sections_And_Bracketed_Currency() {
        const string format = "[$EUR-407]#,##0.00;[Red]-[$EUR-407]#,##0.00";

        SpreadsheetNumberFormatSyntax syntax = SpreadsheetNumberFormatSyntax.Parse(format);

        Assert.Equal(format, string.Concat(syntax.Tokens.Select(token => token.Text)));
        Assert.Equal(2, syntax.SectionCount);
        Assert.Equal("EUR", syntax.CurrencySymbol);
        Assert.True(syntax.UsesGrouping);
    }

    [Fact]
    public void Unterminated_Quoted_Format_Is_Invalid_But_Lossless() {
        const string format = "0.00\"suffix";

        SpreadsheetNumberFormatSyntax syntax = SpreadsheetNumberFormatSyntax.Parse(format);

        Assert.False(syntax.IsValid);
        Assert.Equal(format, string.Concat(syntax.Tokens.Select(token => token.Text)));
    }

    [Theory]
    [InlineData("0,", false, 1)]
    [InlineData("#,##0,", true, 1)]
    [InlineData("0,,", false, 2)]
    public void Parser_Distinguishes_Grouping_From_Thousands_Scaling(
        string format,
        bool usesGrouping,
        int scaleThousands) {
        SpreadsheetNumberFormatSyntax syntax = SpreadsheetNumberFormatSyntax.Parse(format);

        Assert.True(syntax.IsValid);
        Assert.Equal(usesGrouping, syntax.UsesGrouping);
        Assert.Equal(scaleThousands, syntax.ScaleThousands);
        Assert.Equal(format, string.Concat(syntax.Tokens.Select(token => token.Text)));
    }
}