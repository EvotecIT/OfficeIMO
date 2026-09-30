using OfficeIMO.Email;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Markdown;
using System.Net;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Workflows;

public static partial class EmailEvidenceWorkflow {
    private static EmailEvidenceResult Build(string fingerprint, string fingerprintKind, string format, bool graphComplete,
        IEnumerable<(string Id, EmailDocument Document)> sourceMessages, IReadOnlyList<EmailEvidenceThreadLink> links,
        IReadOnlyList<EmailEvidenceMissingParent> missing, List<EmailEvidenceDiagnostic> diagnostics,
        EmailEvidenceOptions options, CancellationToken token,
        Func<IEnumerable<EmailEvidenceDiagnostic>>? finalDiagnostics = null) {
        var html = new StringBuilder("<main><h1>Email evidence</h1><p>Readable projection; signatures and protected content have not been verified or decrypted.</p>");
        var markdown = MarkdownDoc.Create().H1("Email evidence").P("Readable projection; signatures and protected content have not been verified or decrypted.");
        AppendLiteral("Source fingerprint", fingerprintKind + ": " + fingerprint);
        AppendLiteral("Graph coverage", graphComplete ? "Complete within the selected graph scope" : "Incomplete: graph bounds or read failures omitted evidence");
        var messages = new List<EmailEvidenceMessage>();
        foreach (var (id, document) in sourceMessages) {
            token.ThrowIfCancellationRequested();
            if (document.Attachments.Count > 1024) throw new InvalidDataException("The message attachment index exceeds 1024 entries.");
            var text = EmailIndexText.Create(document, new EmailIndexTextOptions {
                PreferPlainText = true, MaxSourceChars = options.MaxBodySourceCharacters, MaxTextChars = options.MaxBodyTextCharacters
            });
            diagnostics.AddRange(text.Diagnostics.Select(MapDiagnostic));
            var attachments = document.Attachments.Select((attachment, index) => new EmailEvidenceAttachment(index,
                ClipNullable(attachment.FileName), ClipNullable(attachment.ContentType), attachment.Length, attachment.IsInline,
                attachment.EmbeddedDocument != null || attachment.MapiAttachMethod == 5, attachment.Content == null ? null :
                    Convert.ToHexString(SHA256.HashData(attachment.Content)).ToLowerInvariant())).ToArray();
            string recipients(EmailRecipientKind kind) => Clip(string.Join(", ", document.Recipients.Where(value => value.Kind == kind)
                .Select(value => Clip(value.Address.ToString())).Take(1000)));
            var message = new EmailEvidenceMessage(id, ClipNullable(document.Subject), ClipNullable(document.From?.ToString()),
                recipients(EmailRecipientKind.To), recipients(EmailRecipientKind.Cc), document.Date, document.ReceivedDate,
                document.Protection.Kind.ToString(), "Unverified", text.SourceKind.ToString(), text.Truncated, attachments);
            messages.Add(message);
            AppendLiteral("Message", message.Subject ?? "Untitled message");
            AppendLiteral("Envelope", $"Item: {id}\nFrom: {message.From}\nTo: {message.To}\nCc: {message.Cc}\nSent: {message.Date:O}\nReceived: {message.ReceivedDate:O}\nProtection: {message.ProtectionKind}\nIntegrity: {message.IntegrityStatus}");
            AppendLiteral("Body", text.FullText);
            if (text.Truncated) AppendLiteral("Body omission", "The configured text limit omitted part of this body.");
            foreach (var attachment in attachments) {
                token.ThrowIfCancellationRequested();
                AppendLiteral("Attachment " + attachment.Index, $"Name: {attachment.Name}\nType: {attachment.ContentType}\nBytes: {attachment.Length}\nInline: {attachment.Inline}\nEmbedded message: {attachment.EmbeddedMessage}\nSHA-256: {attachment.Sha256 ?? "unavailable without deferred payload I/O"}");
            }
        }
        if (finalDiagnostics != null) diagnostics.AddRange(finalDiagnostics());
        foreach (var link in links) AppendLiteral("Thread link", $"{link.ParentId} → {link.ChildId}\n{link.Kind}; {string.Join(", ", link.Reasons)}\nHeuristic: {link.IsHeuristic}");
        foreach (var parent in missing) AppendLiteral("Unresolved parent", $"Child: {parent.ChildId}\nDeclared parent: {parent.ParentMessageId}\nReason: {parent.Reason}");
        foreach (var diagnostic in diagnostics.Take(500)) AppendLiteral("Diagnostic", $"{diagnostic.Code} ({diagnostic.Severity}): {diagnostic.Message}\n{diagnostic.Location}");
        html.Append("</main>");
        string reportHtml = OfficeHtmlDocumentShell.WrapBody(html.ToString(), new OfficeHtmlDocumentOptions { Title = "Email evidence" });
        string reportMarkdown = markdown.ToMarkdown();
        EnsureBounds(reportHtml.Length, reportMarkdown.Length);
        byte[]? pdf = null;
        if (options.IncludePdf) {
            var conversion = HtmlConversionDocument.Parse(reportHtml).ToPdfDocumentResult(new HtmlToPdfOptions {
                MaxInputCharacters = options.MaxReportCharacters, MaxPageCount = options.MaxPdfPages
            }, token);
            pdf = conversion.ToBytes(token);
            diagnostics.AddRange(conversion.Report.Warnings.Select(value => new EmailEvidenceDiagnostic(
                Clip(value.Code), value.Severity.ToString(), Clip(value.Message), ClipNullable(value.Source))));
        }
        var manifest = new EmailEvidenceManifest(1, fingerprint, fingerprintKind, format, graphComplete, messages.ToArray(),
            links, missing, diagnostics.Take(500).ToArray(), diagnostics.Count);
        return new EmailEvidenceResult(reportHtml, reportMarkdown, pdf, manifest, options.MaxOutputBytes);

        void AppendLiteral(string label, string value) {
            token.ThrowIfCancellationRequested();
            html.Append("<section><h2>").Append(WebUtility.HtmlEncode(label)).Append("</h2><div style=\"padding:16px;border:1px solid #cbd5e1;border-radius:8px;overflow-wrap:anywhere;word-break:break-all\">");
            foreach (string line in value.Replace("\r\n", "\n").Replace('\r', '\n').Split('\n')) {
                html.Append("<p style=\"margin:0 0 4px;overflow-wrap:anywhere;word-break:break-all\">")
                    .Append(WebUtility.HtmlEncode(line.Length == 0 ? " " : line)).Append("</p>");
                EnsureBounds(html.Length, 0);
            }
            html.Append("</div></section>");
            markdown.H2(label).Code("text", value);
            EnsureBounds(html.Length, 0);
        }
        void EnsureBounds(int htmlLength, int markdownLength) {
            if (htmlLength > options.MaxReportCharacters || markdownLength > options.MaxReportCharacters)
                throw new InvalidDataException("The email evidence report exceeds MaxReportCharacters.");
        }
    }
}
