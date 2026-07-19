using System;

namespace OfficeIMO.Reader;

/// <summary>Normalizes scalar link and destination values without resolving or fetching targets.</summary>
public sealed class OfficeDocumentLinkNormalizationProcessor : OfficeDocumentProcessorBase {
    /// <summary>Creates the processor.</summary>
    public OfficeDocumentLinkNormalizationProcessor()
        : base("officeimo.reader.normalize-links") {
    }

    /// <inheritdoc />
    public override OfficeDocumentReadResult Process(
        OfficeDocumentReadResult document,
        OfficeDocumentProcessorContext context) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        foreach (OfficeDocumentLink link in OfficeDocumentModelTraversal.Links(document)) {
            context.CancellationToken.ThrowIfCancellationRequested();
            link.Id = (link.Id ?? string.Empty).Trim();
            link.Kind = (link.Kind ?? string.Empty).Trim().ToLowerInvariant();
            link.Uri = TrimOrNull(link.Uri);
            link.DestinationName = TrimOrNull(link.DestinationName);
            link.DestinationMode = TrimOrNull(link.DestinationMode);
            link.NamedAction = TrimOrNull(link.NamedAction);
            link.RemoteFile = TrimOrNull(link.RemoteFile);
            link.RemoteDestinationName = TrimOrNull(link.RemoteDestinationName);
            link.RemoteDestinationMode = TrimOrNull(link.RemoteDestinationMode);
            link.Text = TrimOrNull(link.Text);
        }
        return document;
    }

    private static string? TrimOrNull(string? value) {
        if (value == null) return null;
        string trimmed = value.Trim();
        return trimmed.Length == 0 ? null : trimmed;
    }
}
