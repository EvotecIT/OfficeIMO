namespace OfficeIMO.Email.Data;

/// <summary>Inspects local email-data through its canonical owners without network, certificate discovery or mutation.</summary>
public static class EmailDataInspector {
    /// <summary>Opens an artifact under the supplied read policy, snapshots bounded metadata and closes owned resources.</summary>
    public static EmailDataInspectionReport Inspect(string path, EmailDataInspectionOptions? options = null,
        CancellationToken cancellationToken = default) {
        var policy = options ?? new EmailDataInspectionOptions();
        using var opened = EmailDataArtifact.Open(path, policy.OpenOptions, cancellationToken);
        return Inspect(opened, policy, cancellationToken);
    }

    /// <summary>Snapshots an already opened artifact without disposing it or replacing its owner-level read policy.</summary>
    public static EmailDataInspectionReport Inspect(EmailDataOpenResult opened, EmailDataInspectionOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (opened == null) throw new ArgumentNullException(nameof(opened));
        cancellationToken.ThrowIfCancellationRequested();
        var policy = options ?? new EmailDataInspectionOptions();
        string format = opened.EmailDocument?.Format.ToString() ?? opened.Store?.Format.ToString() ??
            (opened.Calendar != null ? "iCalendar" : opened.Contact != null ? "vCard" : "OAB");
        var report = new EmailDataInspectionReport(opened.Kind, format);
        if (opened.EmailDocument != null) {
            EmailDocument document = opened.EmailDocument;
            report.MessageInspected = true; report.ProtectionKind = document.Protection.Kind.ToString();
            var bodies = new List<EmailDataBodyAlternative>();
            AddBody("PlainText", document.Body.Text, document.Body.TextCharset);
            AddBody("Html", document.Body.Html, document.Body.HtmlCharset); AddBody("Rtf", document.Body.Rtf, null);
            report.Bodies = bodies.AsReadOnly(); report.AttachmentCount = document.Attachments.Count;
            report.Attachments = Array.AsReadOnly(document.Attachments.Take(policy.MaxSamples).Select((attachment, index) => {
                cancellationToken.ThrowIfCancellationRequested();
                return new EmailDataAttachmentMetadata(index, Clip(attachment.FileName), Clip(attachment.ContentType),
                    attachment.Content?.LongLength ?? attachment.ContentSource?.Length ?? attachment.Length,
                    attachment.IsInline, attachment.EmbeddedDocument != null || attachment.MapiAttachMethod == 5 ||
                    string.Equals(attachment.ContentType, "message/rfc822", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(attachment.ContentType, "message/global", StringComparison.OrdinalIgnoreCase),
                    !string.IsNullOrEmpty(attachment.LinkedPath));
            }).ToArray());
            var signatures = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (EmailHeader header in document.Headers.Take(policy.MaxHeadersInspected)) {
                cancellationToken.ThrowIfCancellationRequested();
                if (EmailTransportIntegrity.IsSignatureChain(header.Name)) signatures.Add(header.Name.ToUpperInvariant());
            }
            report.SignatureHeaderNames = Array.AsReadOnly(signatures.OrderBy(name => name, StringComparer.Ordinal).ToArray());
            report.HeaderScanTruncated = document.Headers.Count > policy.MaxHeadersInspected;
            AddDiagnostics(opened.Email!.Diagnostics.Select(d => new EmailDataInspectionDiagnostic(Clip(d.Code)!, d.Severity.ToString())));

            void AddBody(string kind, string? text, string? charset) {
                if (text != null) bodies.Add(new EmailDataBodyAlternative(kind, text.Length, Clip(charset)));
            }
        } else if (opened.Store != null) {
            report.ContainerCount = opened.Store.Folders.Count;
            report.DeclaredItemCount = opened.Store.Folders.Any(folder => !folder.ItemCount.HasValue) ? (long?)null :
                opened.Store.Folders.Sum(folder => (long)folder.ItemCount.GetValueOrDefault());
            report.Containers = Array.AsReadOnly(opened.Store.Folders.Take(policy.MaxSamples).Select(folder => Clip(folder.Name)!).ToArray());
            AddDiagnostics(opened.Store.Diagnostics.Select(d => new EmailDataInspectionDiagnostic(Clip(d.Code)!, d.Severity.ToString())));
        } else if (opened.AddressBook != null) {
            report.ContainerCount = opened.AddressBook.AddressLists.Count;
            report.DeclaredItemCount = opened.AddressBook.DeclaredEntryCount;
            // List identifiers are stable metadata and avoid projecting address-book entries.
            report.Containers = Array.AsReadOnly(opened.AddressBook.AddressLists.Take(policy.MaxSamples).Select(list => Clip(list.Id)!).ToArray());
            AddDiagnostics(opened.AddressBook.Diagnostics.Select(d => new EmailDataInspectionDiagnostic(Clip(d.Code)!, d.Severity.ToString())));
        } else {
            report.ContentLineRootCount = opened.Calendar?.Calendars.Count ?? opened.Contact!.Cards.Count;
            // Parsing qualifies structural syntax. Full semantic validation remains the format owner's explicit operation.
        }
        cancellationToken.ThrowIfCancellationRequested(); return report;

        void AddDiagnostics(IEnumerable<EmailDataInspectionDiagnostic> source) {
            var samples = new List<EmailDataInspectionDiagnostic>();
            foreach (var diagnostic in source) {
                cancellationToken.ThrowIfCancellationRequested(); report.DiagnosticCount++;
                if (samples.Count < policy.MaxSamples) samples.Add(diagnostic);
            }
            report.Diagnostics = samples.AsReadOnly();
        }
        string? Clip(string? text) {
            if (text == null || text.Length <= policy.MaxPreviewCharacters) return text;
            int count = policy.MaxPreviewCharacters;
            if (char.IsHighSurrogate(text[count - 1]) && char.IsLowSurrogate(text[count])) count--;
            return text.Substring(0, count);
        }
    }
}
