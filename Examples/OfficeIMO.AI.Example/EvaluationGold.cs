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
        int factMatches = FactMarkers?.Count(marker => result.Claims.Any(claim => ContainsCompleteMarker(claim.Text, marker))) ?? 0;
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
    private static bool ContainsCompleteMarker(string text, string marker) {
        if (marker.Length == 0) return false;
        for (int start = text.IndexOf(marker, StringComparison.OrdinalIgnoreCase); start >= 0;
            start = text.IndexOf(marker, start + 1, StringComparison.OrdinalIgnoreCase)) {
            int end = start + marker.Length;
            if (start > 0 && IsWordCharacter(text[start - 1]) || end < text.Length && IsWordCharacter(text[end])) continue;
            if (char.IsDigit(marker[0]) && start > 0) {
                char before = text[start - 1];
                if (before is '-' or '+' or '\u2212' or '.'
                    || before == ',' && start > 1 && char.IsDigit(text[start - 2])) continue;
            }
            if (char.IsDigit(marker[^1]) && end + 1 < text.Length
                && text[end] is '.' or ',' or '-' && char.IsDigit(text[end + 1])) continue;
            return true;
        }
        return false;
    }
    private static bool IsWordCharacter(char value) => char.IsLetterOrDigit(value) || value == '_'
        || char.GetUnicodeCategory(value) is System.Globalization.UnicodeCategory.NonSpacingMark or System.Globalization.UnicodeCategory.SpacingCombiningMark;

    private static double Ratio(int value, int total) => total == 0 ? 0 : (double)value / total;
    private static string[] Cells(IReadOnlyList<string> columns, IReadOnlyList<IReadOnlyList<string>> rows) =>
        columns.Select((column, index) => $"header:{index}:{column}").Concat(rows.SelectMany((row, rowIndex) => row.Select((cell, columnIndex) => $"cell:{rowIndex}:{columnIndex}:{cell}"))).ToArray();
}

internal sealed record EvaluationScore(bool Passed, int MatchedFields, int? ExpectedFields, int ReturnedFields,
    double? TableCellPrecision, double? TableCellRecall, double? FactMarkerRecall,
    bool CitationContractSatisfied, bool SemanticReviewRequired);
