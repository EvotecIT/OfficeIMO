namespace OfficeIMO.Email;

internal sealed partial class EmailSemanticSnapshotBuilder {
    private async Task AddAttachments(EmailDocument document, string prefix, int depth,
        bool useAsync, CancellationToken cancellationToken) {
        for (int index = 0; index < document.Attachments.Count; index++) {
            cancellationToken.ThrowIfCancellationRequested();
            EmailAttachment attachment = document.Attachments[index];
            _attachmentCount++;
            string path = string.Concat(prefix, "/attachments/",
                index.ToString("D8", CultureInfo.InvariantCulture));
            int method = attachment.MapiAttachMethod ??
                (attachment.EmbeddedDocument != null ? 5 : attachment.StructuredStorageStreams.Count > 0 ? 6 : 1);
            var diagnostics = new List<EmailDiagnostic>();
            bool hasContent = attachment.Content != null || attachment.ContentSource != null ||
                attachment.EmbeddedDocument != null || attachment.StructuredStorageStreams.Count > 0;
            MsgPropertyBuilder properties = MsgWriter.CreateAttachmentProperties(
                attachment, index, method, diagnostics, path, hasContent, materializedContent: null);
            AddMapiProperties(path + "/mapi", properties.Properties);
            AddValue(path + "/mime-related", attachment.IsMimeRelated);
            AddValue(path + "/content-type-parameters", attachment.ContentTypeParameters);

            if (_options.IncludeAttachmentContent) {
                await AddAttachmentContent(attachment, path, useAsync, cancellationToken)
                    .ConfigureAwait(false);
            } else {
                AddValue(path + "/content-length", attachment.Content?.LongLength ??
                    attachment.ContentSource?.Length ?? attachment.Length);
            }

            int storageIndex = 0;
            foreach (KeyValuePair<string, byte[]> stream in attachment.StructuredStorageStreams
                .OrderBy(item => item.Key, StringComparer.OrdinalIgnoreCase)) {
                string storagePath = string.Concat(path, "/structured-storage/",
                    storageIndex.ToString("D8", CultureInfo.InvariantCulture));
                AddValue(storagePath + "/name", stream.Key);
                AddValue(storagePath + "/content", stream.Value);
                storageIndex++;
            }

            if (_options.Profile == EmailSemanticComparisonProfile.Strict) {
                AddTnefAttributes(path + "/strict/tnef", attachment.TnefAttributes);
            }
            if (attachment.EmbeddedDocument != null) {
                await AddDocument(attachment.EmbeddedDocument, path + "/embedded",
                    depth + 1, useAsync, cancellationToken).ConfigureAwait(false);
            }
        }
    }

    private async Task AddAttachmentContent(EmailAttachment attachment, string path,
        bool useAsync, CancellationToken cancellationToken) {
        long? declaredLength = attachment.Content?.LongLength ?? attachment.ContentSource?.Length;
        if (declaredLength.HasValue && declaredLength.Value > _options.MaxAttachmentBytes) {
            throw new EmailLimitExceededException(
                nameof(EmailSemanticComparisonOptions.MaxAttachmentBytes),
                declaredLength.Value, _options.MaxAttachmentBytes);
        }

        if (attachment.Content == null && attachment.ContentSource == null) {
            AddValue(path + "/content-unavailable", attachment.Length);
            return;
        }

        byte[] digest;
        long bytesRead;
        if (useAsync) {
            using (Stream input = await attachment.OpenContentStreamAsync(cancellationToken)
                .ConfigureAwait(false)) {
                EmailSemanticStreamDigest result = await EmailSemanticValueDigest.ComputeStreamAsync(
                    input, _key, _options.MaxAttachmentBytes, cancellationToken)
                    .ConfigureAwait(false);
                digest = result.Digest;
                bytesRead = result.Length;
            }
        } else {
            using (Stream input = attachment.OpenContentStream()) {
                digest = EmailSemanticValueDigest.ComputeStream(input, _key,
                    _options.MaxAttachmentBytes, cancellationToken, out bytesRead);
            }
        }

        _attachmentBytesHashed = checked(_attachmentBytesHashed + bytesRead);
        if (_attachmentBytesHashed > _options.MaxTotalAttachmentBytes) {
            throw new EmailLimitExceededException(
                nameof(EmailSemanticComparisonOptions.MaxTotalAttachmentBytes),
                _attachmentBytesHashed, _options.MaxTotalAttachmentBytes);
        }
        AddEntry(path + "/content", digest, bytesRead);
    }
}
