namespace OfficeIMO.Email;

/// <summary>Shared policy for MIME transport metadata whose validity depends on serialized bytes.</summary>
internal static class EmailTransportIntegrity {
    internal static bool IsBodySignature(string name) =>
        name.Equals("DKIM-Signature", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("DomainKey-Signature", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("ARC-Message-Signature", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("ARC-Seal", StringComparison.OrdinalIgnoreCase);

    internal static bool IsSignatureChain(string name) => IsBodySignature(name) ||
        name.Equals("ARC-Authentication-Results", StringComparison.OrdinalIgnoreCase);

    internal static bool IsPayloadDependent(string name) =>
        name.Equals("Content-Length", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("Content-MD5", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("Content-Digest", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("Repr-Digest", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("Digest", StringComparison.OrdinalIgnoreCase);

    internal static bool ShouldOmit(string name, EmailWriterOptions options) => IsPayloadDependent(name) ||
        (options.SignatureMutationPolicy == OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures &&
         IsSignatureChain(name));

    internal static void EnsureMutationAllowed(EmailDocument document, OfficeSignatureMutationPolicy policy) {
        if (policy == OfficeSignatureMutationPolicy.BlockSave && document.Headers.Any(header => IsBodySignature(header.Name)))
            throw new InvalidOperationException("Regeneration invalidates transport signatures. Preserve the unchanged source or explicitly choose removal or preservation of invalidated signature markup.");
    }

    internal static void Analyze(EmailDocument root, EmailWriterOptions options, IList<EmailDiagnostic> diagnostics,
        bool writesMime = false) {
        var pending = new Stack<(EmailDocument Document, string Path, int Depth)>();
        var visited = new HashSet<EmailDocument>();
        pending.Push((root, "headers", 0));
        while (pending.Count > 0) {
            var item = pending.Pop();
            if (!visited.Add(item.Document)) continue;
            if (item.Depth > options.MaxNestedMessageDepth)
                throw new EmailLimitExceededException(nameof(options.MaxNestedMessageDepth), item.Depth, options.MaxNestedMessageDepth);
            if (item.Document.Headers.Any(header => IsBodySignature(header.Name))) {
                bool blocked = options.SignatureMutationPolicy == OfficeSignatureMutationPolicy.BlockSave;
                bool removed = options.SignatureMutationPolicy == OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures;
                diagnostics.Add(new EmailDiagnostic("EMAIL_TRANSPORT_SIGNATURE_INVALIDATED",
                    blocked ? "Regeneration would invalidate retained transport signatures; writing was blocked." :
                    removed ? "Transport signature headers are removed from regenerated output." :
                    "Transport signature markup is retained by explicit policy; its validity is not preserved by regeneration.",
                    blocked ? EmailDiagnosticSeverity.Error : EmailDiagnosticSeverity.Warning, item.Path,
                    lossKind: OfficeConversionLossKind.Omission));
            }
            ReportPayloadHeaders(item.Document.Headers, item.Path);
            ReportPayloadHeaders(item.Document.Body.HtmlMimeHeaders, item.Path + "/html");
            for (int index = 0; index < item.Document.Attachments.Count; index++) {
                EmailAttachment attachment = item.Document.Attachments[index];
                ReportPayloadHeaders(attachment.MimeHeaders, item.Path + "/attachment/" + index);
                // MIME cleanup can keep an embedded message's exact payload while rewriting its parent.
                if (writesMime && attachment.EmbeddedDocument != null && MimeWriter.CanPreservePartHeaders(attachment)) continue;
                EmailDocument? child = attachment.EmbeddedDocument;
                if (child != null) pending.Push((child, item.Path + "/attachment/" + index, item.Depth + 1));
            }
            if (!writesMime && OutlookTaskCommunicationAttachmentProjection.GetEmbeddedTaskForWriting(item.Document) is EmailDocument task)
                pending.Push((task, item.Path + "/task/embedded", item.Depth + 1));
        }

        void ReportPayloadHeaders(IEnumerable<EmailHeader> headers, string location) {
            if (headers.Any(header => IsPayloadDependent(header.Name)))
                diagnostics.Add(new EmailDiagnostic("EMAIL_PAYLOAD_METADATA_REMOVED",
                    "Retained length and digest headers are omitted because they describe the original serialized payload.",
                    EmailDiagnosticSeverity.Warning, location, lossKind: OfficeConversionLossKind.Omission));
        }
    }
}
