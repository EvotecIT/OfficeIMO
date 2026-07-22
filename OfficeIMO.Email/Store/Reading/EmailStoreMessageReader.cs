using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal static class EmailStoreMessageReader {
    internal static EmailReadResult Read(byte[] bytes, EmailStoreReaderOptions options,
        CancellationToken cancellationToken, bool? includeAttachmentContent = null) {
        try {
            return new EmailDocumentReader(CreateOptions(options, includeAttachmentContent))
                .Read(bytes, cancellationToken);
        } catch (EmailLimitExceededException exception) {
            throw ConvertLimit(exception);
        }
    }

    internal static EmailReadResult Read(Stream stream, EmailStoreReaderOptions options,
        CancellationToken cancellationToken, bool? includeAttachmentContent = null) {
        try {
            return new EmailDocumentReader(CreateOptions(options, includeAttachmentContent))
                .Read(stream, cancellationToken);
        } catch (EmailLimitExceededException exception) {
            throw ConvertLimit(exception);
        }
    }

    internal static EmailReaderOptions CreateOptions(EmailStoreReaderOptions options,
        bool? includeAttachmentContent = null,
        long? maxDecodedPropertyBytes = null) =>
        new EmailReaderOptions(
            maxInputBytes: options.MaxMessageBytes,
            maxAttachmentBytes: options.MaxAttachmentBytes,
            maxTotalAttachmentBytes: options.MaxTotalAttachmentBytes,
            maxNestedMessageDepth: options.MaxNestedMessageDepth,
            includeAttachmentContent: includeAttachmentContent ?? options.RetainAttachmentContent,
            maxMapiPropertyCount: options.MaxPropertiesPerItem,
            maxDecodedPropertyBytes: maxDecodedPropertyBytes ?? options.MaxDecodedPropertyBytesPerItem);

    private static EmailStoreLimitExceededException ConvertLimit(EmailLimitExceededException exception) {
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
                            : exception.LimitName;
        return new EmailStoreLimitExceededException(name, exception.ActualValue, exception.MaximumValue);
    }
}
