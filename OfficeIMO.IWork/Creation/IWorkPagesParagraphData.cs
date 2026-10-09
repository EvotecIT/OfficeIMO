namespace OfficeIMO.IWork;

/// <summary>An immutable body paragraph in a newly created native Pages document.</summary>
public sealed class IWorkPagesParagraphData {
    internal IWorkPagesParagraphData(IEnumerable<IWorkTextRunData> runs) {
        if (runs == null) throw new ArgumentNullException(nameof(runs));
        IWorkTextRunData[] input = runs.ToArray();
        if (input.Length == 0 || input.Any(run => run == null))
            throw new ArgumentException("At least one non-null character run is required.", nameof(runs));
        if (input.Any(run => run.Text.IndexOf('\n') >= 0))
            throw new ArgumentException("Add each body paragraph separately; character runs cannot contain line breaks.", nameof(runs));
        IWorkTextRunData[] retained = input.Where(run => run.Text.Length > 0).ToArray();
        if (retained.Length == 0) retained = new[] { input[0] };
        Runs = Array.AsReadOnly(retained);
        Text = string.Concat(retained.Select(run => run.Text));
    }

    /// <summary>Gets the paragraph text, without the native paragraph terminator.</summary>
    public string Text { get; }
    /// <summary>Gets the immutable runs in text order.</summary>
    public IReadOnlyList<IWorkTextRunData> Runs { get; }
}
