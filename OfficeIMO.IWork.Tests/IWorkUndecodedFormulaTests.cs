using OfficeIMO.IWork;

namespace OfficeIMO.IWork.Tests;

public sealed partial class IWorkBoundaryTests {
    public static IEnumerable<object[]> UndecodedFormulaCases() {
        foreach (IWorkDocumentKind kind in new[] { IWorkDocumentKind.Pages,
                IWorkDocumentKind.Numbers, IWorkDocumentKind.Keynote }) {
            foreach (string failure in new[] { "unsupportedFields", "truncatedValue", "conflictingCache",
                    "nonFiniteCache", "invalidBoolean", "unknownType" }) {
                foreach (bool visual in new[] { false, true }) yield return new object[] { kind, failure, visual };
            }
        }
    }

    [Theory]
    [MemberData(nameof(UndecodedFormulaCases))]
    public void Supported_header_formula_declaration_survives_cell_decode_failure(
        IWorkDocumentKind kind, string failure, bool visual) {
        using MemoryStream package = TableDependencyPackage(kind, Message(),
            cellPayload: UndecodedFormulaCell(failure));
        IWorkSourceDocument source = IWorkSourceDocument.Open(package, kind);
        (IWorkTable table, _) = ReadSelectedRichTable(source, kind);
        IWorkTableCell cell = Assert.Single(table.Cells);
        Assert.Equal(IWorkCellKind.Error, cell.Kind);
        Assert.True(cell.SourceFormulaIsDeclared);
        Assert.True(cell.HasDecodeError);
        Assert.NotNull(cell.Error);
        Assert.Null(cell.Value);
        Assert.Null(cell.Formula);
        package.Position = 0;
        IWorkConversionReport report = ConvertUnitReport(package, kind, visual,
            new IWorkReadOptions { PreserveSourceRecords = false });
        Assert.Empty(report.PreservedRecords);
        Assert.Equal(visual ? IWorkProjectionKind.VisualFallback : IWorkProjectionKind.EditableReconstruction,
            report.ProjectionKind);
        IWorkSourceCellIssue issue = Assert.Single(report.SourceCellIssues);
        Assert.Equal(10ul, issue.TableIdentity!.RecordIdentifier);
        Assert.Equal((1, 1), (issue.Row, issue.Column));
        Assert.Equal(cell.Error, issue.Message);
        IWorkFormulaCellStatus assessment = Assert.Single(report.FormulaCells);
        Assert.Equal(10ul, assessment.TableIdentity!.RecordIdentifier);
        Assert.Equal(1, assessment.Row); Assert.Equal(1, assessment.Column);
        Assert.False(assessment.ExpressionIsAssessed);
        Assert.False(assessment.ExpressionIsComplete);
        Assert.Equal(IWorkFormulaCacheStatus.Unassessed, assessment.CacheStatus);
        Assert.Null(assessment.CachedValueKind);
        Assert.Equal(1, report.FormulaSummary.TotalCount);
        Assert.Equal(1, report.FormulaSummary.UnassessedExpressionCount);
        Assert.Equal(1, report.FormulaSummary.UnassessedCacheCount);
        Assert.Equal(0, report.FormulaSummary.IncompleteExpressionCount);
        Assert.Equal(0, report.FormulaSummary.MissingCacheCount);
        Assert.Equal(OfficeConversionLossKind.Unassessed, Assert.Single(report.FidelityDiagnostics,
            diagnostic => diagnostic.Code == "IWORK_FORMULA_CELLS_UNASSESSED").LossKind);
        Assert.Throws<NotSupportedException>(() => ((IList<IWorkFormulaCellStatus>)report.FormulaCells).Clear());
    }

    [Theory]
    [InlineData(IWorkDocumentKind.Pages)]
    [InlineData(IWorkDocumentKind.Numbers)]
    [InlineData(IWorkDocumentKind.Keynote)]
    public void Unknown_headers_and_nonformula_errors_do_not_invent_formula_declarations(IWorkDocumentKind kind) {
        foreach (string failure in new[] { "unsupportedVersion", "truncatedHeader", "nonFormula" }) {
            byte[] cell = UndecodedFormulaCell("truncatedValue");
            if (failure == "unsupportedVersion") cell[0] = 4;
            else if (failure == "truncatedHeader") cell = cell.Take(11).ToArray();
            else WriteUInt32(cell, 8, 1u << 1);
            using MemoryStream package = TableDependencyPackage(kind, Message(), cellPayload: cell);
            (IWorkTable table, _) = ReadSelectedRichTable(IWorkSourceDocument.Open(package, kind), kind);
            Assert.False(Assert.Single(table.Cells).SourceFormulaIsDeclared);
            package.Position = 0;
            IWorkConversionReport report = ConvertUnitReport(package, kind);
            Assert.Empty(report.FormulaCells);
            Assert.Equal(0, report.FormulaSummary.TotalCount);
        }
    }

    private static byte[] UndecodedFormulaCell(string failure) {
        byte[] cell = new byte[failure == "truncatedValue" ? 12 : failure == "conflictingCache" ? 28 : 24];
        cell[0] = 5; cell[1] = failure == "invalidBoolean" ? (byte)6 : failure == "unknownType" ? (byte)99 : (byte)2;
        uint flags = (1u << 1) | (1u << 9);
        if (failure == "unsupportedFields") flags |= 1u << 31;
        if (failure == "conflictingCache") flags |= 1u << 3;
        WriteUInt32(cell, 8, flags);
        if (cell.Length > 12) Buffer.BlockCopy(BitConverter.GetBytes(
            failure == "nonFiniteCache" ? double.NaN : failure == "invalidBoolean" ? 2d : 42d), 0, cell, 12, 8);
        return cell;
    }
}
