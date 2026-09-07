using OfficeIMO.AI;

// Gold values are declared before inference, independently of prompts and model responses.
internal sealed record EvaluationFieldGold(string Name, OfficeAiFieldStatus Status, string? Value = null);
internal sealed record EvaluationTableGold(IReadOnlyList<string> Columns, IReadOnlyList<IReadOnlyList<string>> Rows);
internal sealed record EvaluationGold(IReadOnlyList<EvaluationFieldGold>? Fields = null,
    EvaluationTableGold? Table = null, IReadOnlyList<string>? FactMarkers = null,
    OfficeAiResultStatus? Status = OfficeAiResultStatus.Completed, bool RequireSynthesis = false) {
    public EvaluationScore Score(OfficeAiResult result) {
        int fieldMatches = Fields?.Count(expected => result.Fields.Any(actual => actual.Name == expected.Name
            && actual.Status == expected.Status && actual.NormalizedValue == expected.Value)) ?? 0;
        string[] expectedCells = Table is null ? Array.Empty<string>() : Cells(Table.Columns, Table.Rows);
        string[] actualCells = result.Tables.SelectMany(table => Cells(table.Table.Columns, table.Table.Rows)).ToArray();
        // Multiset matching penalizes extra tables and duplicate cells; positions preserve row/column relationships.
        var available = expectedCells.GroupBy(value => value, StringComparer.Ordinal).ToDictionary(group => group.Key, group => group.Count(), StringComparer.Ordinal);
        int cellMatches = 0;
        foreach (string cell in actualCells)
            if (available.TryGetValue(cell, out int count) && count > 0) { cellMatches++; available[cell] = count - 1; }
        string claims = string.Join(" ", result.Claims.Select(claim => claim.Text));
        int factMatches = FactMarkers?.Count(marker => claims.Contains(marker, StringComparison.OrdinalIgnoreCase)) ?? 0;
        bool fieldsCorrect = Fields is null || fieldMatches == Fields.Count && result.Fields.Count == Fields.Count;
        bool tableCorrect = Table is null || cellMatches == expectedCells.Length && cellMatches == actualCells.Length;
        bool markersCorrect = FactMarkers is null || factMatches == FactMarkers.Count;
        bool stateCorrect = !Status.HasValue || result.Status == Status;
        bool synthesisCorrect = !RequireSynthesis || result.SynthesisStatus == OfficeAiSynthesisStatus.Completed;
        return new(fieldsCorrect && tableCorrect && markersCorrect && stateCorrect && synthesisCorrect,
            fieldMatches, Fields?.Count, result.Fields.Count,
            Table is null ? null : Ratio(cellMatches, actualCells.Length), Table is null ? null : Ratio(cellMatches, expectedCells.Length),
            FactMarkers is null ? null : Ratio(factMatches, FactMarkers.Count),
            result.Claims.SelectMany(claim => claim.Citations).All(citation => citation.QuoteMatched || citation.Quote is null),
            result.RequiresReview);
    }
    private static double Ratio(int value, int total) => total == 0 ? 0 : (double)value / total;
    private static string[] Cells(IReadOnlyList<string> columns, IReadOnlyList<IReadOnlyList<string>> rows) =>
        columns.Select((column, index) => $"header:{index}:{column}").Concat(rows.SelectMany((row, rowIndex) => row.Select((cell, columnIndex) => $"cell:{rowIndex}:{columnIndex}:{cell}"))).ToArray();
}

internal sealed record EvaluationScore(bool Passed, int MatchedFields, int? ExpectedFields, int ReturnedFields,
    double? TableCellPrecision, double? TableCellRecall, double? FactMarkerRecall,
    bool CitationContractSatisfied, bool SemanticReviewRequired);
