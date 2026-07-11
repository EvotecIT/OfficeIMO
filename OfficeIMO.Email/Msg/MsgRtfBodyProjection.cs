using OfficeIMO.Rtf;
using OfficeIMO.Rtf.Diagnostics;

namespace OfficeIMO.Email;

/// <summary>Projects Outlook HTML encapsulation through the owning OfficeIMO.Rtf engine.</summary>
internal static class MsgRtfBodyProjection {
    internal static string? TryGetEncapsulatedHtml(string? rtf, MsgParserState state, string location) {
        if (string.IsNullOrEmpty(rtf)) return null;

        try {
            RtfReadOptions options = RtfReadOptions.CreateUntrustedProfile();
            int maximum = checked((int)Math.Min(state.Options.MaxDecodedPropertyBytes, int.MaxValue));
            options.MaxInputBytes = maximum;
            options.MaxInputCharacters = maximum;
            options.MaxTextCharacters = maximum;
            RtfReadResult result = RtfDocument.Read(rtf!, options, state.CancellationToken);
            AddDiagnostics(result.Diagnostics, state.Diagnostics, location);
            string? html = result.Document.HtmlEncapsulation?.Html;
            return string.IsNullOrWhiteSpace(html) ? null : html;
        } catch (RtfReadLimitException ex) {
            state.Diagnostics.Add(new EmailDiagnostic(
                "EMAIL_MSG_RTF_LIMIT_EXCEEDED",
                ex.Message,
                EmailDiagnosticSeverity.Warning,
                location));
        } catch (Exception ex) when (ex is InvalidDataException || ex is FormatException || ex is ArgumentException) {
            state.Diagnostics.Add(new EmailDiagnostic(
                "EMAIL_MSG_RTF_PROJECTION_FAILED",
                ex.Message,
                EmailDiagnosticSeverity.Warning,
                location));
        }
        return null;
    }

    private static void AddDiagnostics(
        IReadOnlyList<RtfDiagnostic> rtfDiagnostics,
        IList<EmailDiagnostic> emailDiagnostics,
        string location) {
        for (int index = 0; index < rtfDiagnostics.Count; index++) {
            RtfDiagnostic diagnostic = rtfDiagnostics[index];
            if (diagnostic.Severity == RtfDiagnosticSeverity.Info) continue;
            emailDiagnostics.Add(new EmailDiagnostic(
                string.Concat("EMAIL_MSG_RTF_", diagnostic.Code),
                diagnostic.Message,
                diagnostic.Severity == RtfDiagnosticSeverity.Error
                    ? EmailDiagnosticSeverity.Error
                    : EmailDiagnosticSeverity.Warning,
                string.Concat(location, "@", diagnostic.Position.ToString(CultureInfo.InvariantCulture))));
        }
    }
}
