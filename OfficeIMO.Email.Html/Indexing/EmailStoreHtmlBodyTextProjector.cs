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
        EmailIndexTextResult projection = EmailIndexText.Create(view, new EmailIndexTextOptions {
            MaxTextChars = maxCharacters, InspectContentSafety = true, ConcealedTextPolicy = _policy
        });
        cancellationToken.ThrowIfCancellationRequested();
        if (projection.ContentSafety != null) _inspectionObserver?.Invoke(projection.ContentSafety);
        return projection.SelectedText;
    }
}
