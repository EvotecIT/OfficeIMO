using OfficeIMO.Email.Store;

namespace OfficeIMO.Email;

/// <summary>Network-free semantic-search body projection through the shared email/HTML owner.</summary>
public sealed class EmailStoreHtmlBodyTextProjector : IEmailStoreBodyTextProjector {
    private readonly EmailConcealedTextPolicy _policy;
    private readonly Action<EmailBodyContentSafetyReport>? _inspectionObserver;

    /// <summary>Creates a bounded projector. The optional observer receives advisory evidence without source text.</summary>
    public EmailStoreHtmlBodyTextProjector(EmailConcealedTextPolicy policy = EmailConcealedTextPolicy.Preserve,
        Action<EmailBodyContentSafetyReport>? inspectionObserver = null) {
        if (!Enum.IsDefined(typeof(EmailConcealedTextPolicy), policy)) throw new ArgumentOutOfRangeException(nameof(policy));
        _policy = policy;
        _inspectionObserver = inspectionObserver;
    }

    /// <inheritdoc />
    public string Identity => "OfficeIMO.Email.Html.body-text.v1:" + _policy;

    /// <inheritdoc />
    public string Project(EmailDocument document, EmailStoreContentSearchFields bodyField, int maxCharacters,
        CancellationToken cancellationToken) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (maxCharacters <= 0) throw new ArgumentOutOfRangeException(nameof(maxCharacters));
        cancellationToken.ThrowIfCancellationRequested();
        var view = new EmailDocument();
        switch (bodyField) {
            case EmailStoreContentSearchFields.TextBody: view.Body.Text = document.Body.Text; break;
            case EmailStoreContentSearchFields.HtmlBody: view.Body.Html = document.Body.Html; break;
            case EmailStoreContentSearchFields.RtfBody: view.Body.Rtf = document.Body.Rtf; break;
            default: throw new ArgumentOutOfRangeException(nameof(bodyField));
        }
        if (string.IsNullOrEmpty(view.Body.Text) && string.IsNullOrEmpty(view.Body.Html) && string.IsNullOrEmpty(view.Body.Rtf)) return string.Empty;
        const int maximumSourceCharacters = 2 * 1024 * 1024;
        if ((view.Body.Text ?? view.Body.Html ?? view.Body.Rtf ?? string.Empty).Length > maximumSourceCharacters) {
            _inspectionObserver?.Invoke(new EmailBodyContentSafetyReport { InspectionStatus = "BodyLimitExceeded" });
            // Use the store's recoverable item-error contract rather than rejecting the whole query.
            throw new InvalidDataException("The selected email body exceeds the body projector's 2 MiB character limit.");
        }
        EmailIndexTextResult projection = EmailIndexText.Create(view, new EmailIndexTextOptions {
            MaxSourceChars = maximumSourceCharacters, MaxTextChars = maxCharacters,
            InspectContentSafety = true, ConcealedTextPolicy = _policy
        });
        cancellationToken.ThrowIfCancellationRequested();
        if (projection.ContentSafety != null) _inspectionObserver?.Invoke(projection.ContentSafety);
        return projection.SelectedText;
    }
}
