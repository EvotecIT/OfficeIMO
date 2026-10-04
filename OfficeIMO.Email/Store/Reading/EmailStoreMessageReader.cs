using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal static class EmailStoreMessageReader {
    internal static EmailReadResult Read(byte[] bytes, EmailStoreReaderOptions options,
        CancellationToken cancellationToken, bool? includeAttachmentContent = null, long? maxDecodedPropertyBytes = null,
        bool includeEmbeddedMessages = true) {
        try {
            return new EmailDocumentReader(CreateOptions(options, includeAttachmentContent, maxDecodedPropertyBytes, includeEmbeddedMessages))
                .Read(bytes, cancellationToken);
        } catch (EmailLimitExceededException exception) {
            throw ConvertLimit(exception);
        }
    }

    internal static EmailReadResult Read(Stream stream, EmailStoreReaderOptions options,
        CancellationToken cancellationToken, bool? includeAttachmentContent = null, long? maxDecodedPropertyBytes = null,
        bool includeEmbeddedMessages = true, bool preferStreamingAttachmentContent = false) {
        try {
            var reader = new EmailDocumentReader(CreateOptions(options, includeAttachmentContent, maxDecodedPropertyBytes, includeEmbeddedMessages));
            return preferStreamingAttachmentContent || !(includeAttachmentContent ?? options.RetainAttachmentContent)
                ? reader.ReadStreaming(stream, cancellationToken: cancellationToken)
                : reader.Read(stream, cancellationToken);
        } catch (EmailLimitExceededException exception) {
            throw ConvertLimit(exception);
        }
    }

    internal static EmailReaderOptions CreateOptions(EmailStoreReaderOptions options,
        bool? includeAttachmentContent = null,
        long? maxDecodedPropertyBytes = null,
        bool includeEmbeddedMessages = true) =>
        new EmailReaderOptions(
            maxInputBytes: options.MaxMessageBytes,
            maxAttachmentBytes: options.MaxAttachmentBytes,
            maxTotalAttachmentBytes: options.MaxTotalAttachmentBytes,
            maxNestedMessageDepth: options.MaxNestedMessageDepth,
            includeAttachmentContent: includeAttachmentContent ?? options.RetainAttachmentContent,
            maxMapiPropertyCount: options.MaxPropertiesPerItem,
            maxDecodedPropertyBytes: Math.Min(maxDecodedPropertyBytes ?? options.MaxDecodedPropertyBytesPerItem, options.MaxDecodedPropertyBytesPerItem),
            maxAttachmentCount: options.MaxAttachmentsPerItem,
            includeEmbeddedMessages: includeEmbeddedMessages);

    internal static EmailStoreLimitExceededException ConvertLimit(EmailLimitExceededException exception) {
        string name = exception.LimitName == nameof(EmailReaderOptions.MaxInputBytes)
            ? nameof(EmailStoreReaderOptions.MaxMessageBytes)
            : exception.LimitName == nameof(EmailReaderOptions.MaxAttachmentBytes)
                ? nameof(EmailStoreReaderOptions.MaxAttachmentBytes)
                : exception.LimitName == nameof(EmailReaderOptions.MaxTotalAttachmentBytes)
                    ? nameof(EmailStoreReaderOptions.MaxTotalAttachmentBytes)
                    : exception.LimitName == nameof(EmailReaderOptions.MaxMapiPropertyCount)
                        ? nameof(EmailStoreReaderOptions.MaxPropertiesPerItem)
                        : exception.LimitName == nameof(EmailReaderOptions.MaxDecodedPropertyBytes)
                            ? nameof(EmailStoreReaderOptions.MaxDecodedPropertyBytesPerItem)
                            : exception.LimitName == nameof(EmailReaderOptions.MaxAttachmentCount)
                                ? nameof(EmailStoreReaderOptions.MaxAttachmentsPerItem)
                            : exception.LimitName;
        return new EmailStoreLimitExceededException(name, exception.ActualValue, exception.MaximumValue);
    }
}
