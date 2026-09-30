using OfficeIMO.AI;

internal sealed record EvaluationCase(string Id, string Extension, byte[] Source, OfficeAiRequest Request, bool Images,
    string Expected, EvaluationGold Gold) {
    public string Split { get; init; } = "development";
}

