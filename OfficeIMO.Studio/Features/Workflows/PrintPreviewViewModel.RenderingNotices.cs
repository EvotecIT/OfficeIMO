using System.Collections.ObjectModel;

namespace OfficeIMO.Studio.Features.Workflows;

/// <summary>Localized print-review guidance with the unchanged engine detail available separately.</summary>
public sealed record PrintRenderingNotice(string Code, string Message, string Details);

public sealed partial class PrintPreviewViewModel {
    public ObservableCollection<PrintRenderingNotice> RenderingNotices { get; } = new();
    public bool HasRenderingNotices => RenderingNotices.Count > 0;

    private void SetRenderingNotices(IReadOnlyList<string> diagnostics) {
        RenderingNotices.Clear();
        foreach (string diagnostic in diagnostics) {
            int separator = diagnostic.IndexOf(':');
            string candidate = separator > 0 ? diagnostic[..separator] : string.Empty;
            // PDF render diagnostics prefix their stable code with "render.". An unstructured
            // engine message remains technical detail rather than becoming an English UI label.
            string code = candidate.StartsWith("render.", StringComparison.Ordinal) &&
                candidate.All(character => char.IsAsciiLetterOrDigit(character) || character is '.' or '-' or '_')
                ? candidate : string.Empty;
            string key = code == "render.resource.font-substitution" ? "Rendering.FontSubstitution" : "Rendering.Limitation";
            RenderingNotices.Add(new(code, T(key), diagnostic));
        }
        OnPropertyChanged(nameof(HasRenderingNotices));
    }
}
