using System.Security.Cryptography;
using System.Text;
using MimeKit;
#if NET10_0_OR_GREATER
using OneMailApi = RecursiveAI.OneMail;
#endif
using OfficeIMO.Benchmarks;

namespace OfficeIMO.Email.Benchmarks.Comparisons;

internal sealed record EmailMimeAttachmentObservation(string FileName, int Length, string Sha256);

internal sealed record EmailMimeObservation(
    string Subject,
    string From,
    IReadOnlyList<string> To,
    IReadOnlyList<string> Cc,
    string TextBody,
    string HtmlBody,
    IReadOnlyList<EmailMimeAttachmentObservation> Attachments);

internal sealed record EmailMimeComparisonReport(
    string Scale,
    int InputBytes,
    int OfficeIMOOutputBytes,
    int MimeKitOutputBytes,
    int AttachmentCount,
    int DecodedAttachmentBytes);

internal static class EmailMimeComparisonValidation {
    internal static IReadOnlyList<EmailMimeComparisonReport> ValidateAll() =>
        EmailMimeComparisonCorpus.ScaleNames.Select(Validate).ToArray();

    internal static EmailMimeComparisonReport Validate(string scaleName) {
        EmailMimeBenchmarkScale scale = EmailMimeComparisonCorpus.Get(scaleName);
        EmailDocument officeModel = EmailMimeComparisonCorpus.CreateOfficeDocument(scale);
        using MimeMessage mimeModel = EmailMimeComparisonCorpus.CreateMimeMessage(scale);
        byte[] mimeOutput = EmailMimeComparisonCorpus.WriteMimeKit(mimeModel);
        byte[] officeOutput = officeModel.ToBytes(EmailFileFormat.Eml);

        EmailMimeObservation expected = ObserveMimeKit(mimeOutput);
        RequireEqual(scaleName, "canonical input", expected, ObserveOffice(mimeOutput));
        RequireEqual(scaleName, "OfficeIMO output through OfficeIMO", expected, ObserveOffice(officeOutput));
        RequireEqual(scaleName, "OfficeIMO output through MimeKit", expected, ObserveMimeKit(officeOutput));
#if NET10_0_OR_GREATER
        byte[] oneMailOutput = EmailMimeComparisonCorpus.WriteOneMail(EmailMimeComparisonCorpus.CreateOneMailMessage(scale));
        RequireEqual(scaleName, "canonical input through OneMail", expected, ObserveOneMail(mimeOutput));
        RequireEqual(scaleName, "OneMail output through OneMail", expected, ObserveOneMail(oneMailOutput));
        RequireEqual(scaleName, "OneMail output through MimeKit", expected, ObserveMimeKit(oneMailOutput));
        RequireEqual(scaleName, "OneMail output through OfficeIMO", expected, ObserveOffice(oneMailOutput));
#endif

        return new EmailMimeComparisonReport(
            scale.Name,
            mimeOutput.Length,
            officeOutput.Length,
            mimeOutput.Length,
            expected.Attachments.Count,
            expected.Attachments.Sum(attachment => attachment.Length));
    }

    /// <summary>
    /// Compares what each reader recovered from one library's lane with the corpus' expected headers, bodies and
    /// decoded attachments, records the result for <paramref name="library"/> and throws on a mismatch.
    /// </summary>
    internal static void RecordLane(
        string library,
        string scaleName,
        EmailMimeObservation expected,
        params (string Reader, EmailMimeObservation Observation)[] observations) {
        string detail = $"Scale={scaleName}; subject, From, {expected.To.Count} To, {expected.Cc.Count} Cc, " +
            $"text body {expected.TextBody.Length} chars, HTML body {expected.HtmlBody.Length} chars, " +
            $"{expected.Attachments.Count} attachments ({expected.Attachments.Sum(attachment => attachment.Length)} bytes, SHA-256)";
        foreach ((string reader, EmailMimeObservation observation) in observations) {
            if (!Equivalent(expected, observation)) {
                string message = $"Email MIME content read by {reader} does not match the expected message: {Describe(expected, observation)}.";
                BenchmarkOutputValidationLog.Failed(library, detail + "; " + message);
                throw new InvalidOperationException(message);
            }
        }

        BenchmarkOutputValidationLog.Passed(
            library,
            detail + "; matched through " + string.Join(" and ", observations.Select(observation => observation.Reader)));
    }

    private static string Describe(EmailMimeObservation expected, EmailMimeObservation actual) {
        var differences = new List<string>();
        if (expected.Subject != actual.Subject) differences.Add("subject");
        if (expected.From != actual.From) differences.Add("From");
        if (!expected.To.SequenceEqual(actual.To, StringComparer.Ordinal)) differences.Add("To");
        if (!expected.Cc.SequenceEqual(actual.Cc, StringComparer.Ordinal)) differences.Add("Cc");
        if (expected.TextBody != actual.TextBody) differences.Add($"text body ({actual.TextBody.Length} of {expected.TextBody.Length} chars)");
        if (expected.HtmlBody != actual.HtmlBody) differences.Add($"HTML body ({actual.HtmlBody.Length} of {expected.HtmlBody.Length} chars)");
        if (!expected.Attachments.SequenceEqual(actual.Attachments)) differences.Add($"attachments ({actual.Attachments.Count} of {expected.Attachments.Count})");
        return string.Join(", ", differences);
    }

    internal static int ConsumeOffice(byte[] input) {
        EmailDocument document = EmailDocument.Load(input);
        int checksum = (document.Subject?.Length ?? 0) +
                       (document.Body.Text?.Length ?? 0) +
                       (document.Body.Html?.Length ?? 0) +
                       document.Recipients.Count;
        foreach (EmailAttachment attachment in document.Attachments) {
            byte[] content = ReadAttachment(attachment);
            checksum = ConsumeBytes(checksum, content);
        }
        return checksum;
    }

    internal static int ConsumeMimeKit(byte[] input) {
        using var stream = new MemoryStream(input, writable: false);
        using MimeMessage message = MimeMessage.Load(stream);
        int checksum = (message.Subject?.Length ?? 0) +
                       (message.TextBody?.Length ?? 0) +
                       (message.HtmlBody?.Length ?? 0) +
                       message.To.Count + message.Cc.Count;
        foreach (MimeEntity entity in message.Attachments) {
            if (entity is not MimePart part) continue;
            using var decoded = new MemoryStream();
            (part.Content ?? throw new InvalidDataException("MimeKit attachment content is missing."))
                .DecodeTo(decoded);
            checksum = ConsumeBytes(checksum, decoded.ToArray());
        }
        return checksum;
    }

#if NET10_0_OR_GREATER
    internal static int ConsumeOneMail(byte[] input) {
        OneMailApi.IMailItem item = OneMailApi.Mail.Parse(input);
        int checksum = (item.Subject?.Length ?? 0) +
                       (item.TextBody?.ReadText().Length ?? 0) +
                       (item.HtmlBody?.ReadText().Length ?? 0) +
                       item.ToRecipients.Count + item.Cc.Count;
        foreach (OneMailApi.IMailAttachment attachment in item.Attachments) {
            checksum = ConsumeBytes(checksum, ReadOneMailAttachment(attachment));
        }
        return checksum;
    }

    internal static EmailMimeObservation ObserveOneMail(byte[] input) {
        OneMailApi.IMailItem item = OneMailApi.Mail.Parse(input);
        return new EmailMimeObservation(
            item.Subject ?? string.Empty,
            NormalizeAddress(item.From.Count > 0 ? item.From[0].Address : null),
            item.ToRecipients.Select(address => NormalizeAddress(address.Address)).ToArray(),
            item.Cc.Select(address => NormalizeAddress(address.Address)).ToArray(),
            NormalizeBody(item.TextBody?.ReadText()),
            NormalizeBody(item.HtmlBody?.ReadText()),
            item.Attachments.Select(attachment => {
                byte[] content = ReadOneMailAttachment(attachment);
                return new EmailMimeAttachmentObservation(
                    attachment.FileName ?? string.Empty,
                    content.Length,
                    Convert.ToHexString(SHA256.HashData(content)));
            }).ToArray());
    }

    internal static EmailMimeRetainedProjection RetainOneMail(byte[] input) {
        OneMailApi.IMailItem item = OneMailApi.Mail.Parse(input);
        return new EmailMimeRetainedProjection(
            item,
            item.Subject ?? string.Empty,
            NormalizeBody(item.TextBody?.ReadText()),
            NormalizeBody(item.HtmlBody?.ReadText()),
            item.Attachments.Select(ReadOneMailAttachment).ToArray());
    }

    private static byte[] ReadOneMailAttachment(OneMailApi.IMailAttachment attachment) {
        using Stream decoded = attachment.Open();
        using var output = new MemoryStream();
        decoded.CopyTo(output);
        return output.ToArray();
    }

#endif
    internal static EmailMimeObservation ObserveOffice(byte[] input) {
        EmailDocument document = EmailDocument.Load(input);
        return new EmailMimeObservation(
            document.Subject ?? string.Empty,
            NormalizeAddress(document.From?.Address),
            Addresses(document, EmailRecipientKind.To),
            Addresses(document, EmailRecipientKind.Cc),
            NormalizeBody(document.Body.Text),
            NormalizeBody(document.Body.Html),
            document.Attachments.Select(attachment => {
                byte[] content = ReadAttachment(attachment);
                return new EmailMimeAttachmentObservation(
                    attachment.FileName ?? string.Empty,
                    content.Length,
                    Convert.ToHexString(SHA256.HashData(content)));
            }).ToArray());
    }

    internal static EmailMimeObservation ObserveMimeKit(byte[] input) {
        using var stream = new MemoryStream(input, writable: false);
        using MimeMessage message = MimeMessage.Load(stream);
        return new EmailMimeObservation(
            message.Subject ?? string.Empty,
            NormalizeAddress(message.From.Mailboxes.FirstOrDefault()?.Address),
            message.To.Mailboxes.Select(mailbox => NormalizeAddress(mailbox.Address)).ToArray(),
            message.Cc.Mailboxes.Select(mailbox => NormalizeAddress(mailbox.Address)).ToArray(),
            NormalizeBody(message.TextBody),
            NormalizeBody(message.HtmlBody),
            message.Attachments.OfType<MimePart>().Select(part => {
                using var decoded = new MemoryStream();
                (part.Content ?? throw new InvalidDataException("MimeKit attachment content is missing."))
                    .DecodeTo(decoded);
                byte[] content = decoded.ToArray();
                return new EmailMimeAttachmentObservation(
                    part.FileName ?? string.Empty,
                    content.Length,
                    Convert.ToHexString(SHA256.HashData(content)));
            }).ToArray());
    }

    private static IReadOnlyList<string> Addresses(EmailDocument document, EmailRecipientKind kind) =>
        document.Recipients
            .Where(recipient => recipient.Kind == kind)
            .Select(recipient => NormalizeAddress(recipient.Address.Address))
            .ToArray();

    private static byte[] ReadAttachment(EmailAttachment attachment) {
        if (attachment.Content != null) return attachment.Content;
        using Stream stream = attachment.OpenContentStream();
        using var output = new MemoryStream();
        stream.CopyTo(output);
        return output.ToArray();
    }

    private static int ConsumeBytes(int checksum, byte[] content) {
        unchecked {
            for (int index = 0; index < content.Length; index++) checksum = checksum * 31 + content[index];
        }
        return checksum;
    }

    private static string NormalizeAddress(string? value) => (value ?? string.Empty).Trim().ToLowerInvariant();

    internal static string NormalizeBody(string? value) =>
        (value ?? string.Empty).Replace("\r\n", "\n", StringComparison.Ordinal).Replace('\r', '\n').TrimEnd();

    private static void RequireEqual(
        string scale,
        string path,
        EmailMimeObservation expected,
        EmailMimeObservation actual) {
        if (!Equivalent(expected, actual)) {
            throw new InvalidOperationException($"Email MIME semantic validation failed for {scale}/{path}.");
        }
    }

    private static bool Equivalent(EmailMimeObservation left, EmailMimeObservation right) =>
        left.Subject == right.Subject &&
        left.From == right.From &&
        left.To.SequenceEqual(right.To, StringComparer.Ordinal) &&
        left.Cc.SequenceEqual(right.Cc, StringComparer.Ordinal) &&
        left.TextBody == right.TextBody &&
        left.HtmlBody == right.HtmlBody &&
        left.Attachments.SequenceEqual(right.Attachments);

    internal static EmailMimeRetainedProjection RetainOffice(byte[] input) {
        EmailDocument document = EmailDocument.Load(input);
        return new EmailMimeRetainedProjection(
            document,
            document.Subject ?? string.Empty,
            NormalizeBody(document.Body.Text),
            NormalizeBody(document.Body.Html),
            document.Attachments.Select(ReadAttachment).ToArray());
    }

    internal static EmailMimeRetainedProjection RetainMimeKit(byte[] input) {
        using var stream = new MemoryStream(input, writable: false);
        MimeMessage message = MimeMessage.Load(stream);
        return new EmailMimeRetainedProjection(
            message,
            message.Subject ?? string.Empty,
            NormalizeBody(message.TextBody),
            NormalizeBody(message.HtmlBody),
            message.Attachments.OfType<MimePart>().Select(part => {
                using var decoded = new MemoryStream();
                (part.Content ?? throw new InvalidDataException("MimeKit attachment content is missing."))
                    .DecodeTo(decoded);
                return decoded.ToArray();
            }).ToArray());
    }
}

internal sealed class EmailMimeRetainedProjection : IDisposable {
    internal EmailMimeRetainedProjection(
        object source,
        string subject,
        string textBody,
        string htmlBody,
        IReadOnlyList<byte[]> attachments) {
        Source = source;
        Subject = subject;
        TextBody = textBody;
        HtmlBody = htmlBody;
        Attachments = attachments;
    }

    internal object Source { get; }
    internal string Subject { get; }
    internal string TextBody { get; }
    internal string HtmlBody { get; }
    internal IReadOnlyList<byte[]> Attachments { get; }

    public void Dispose() {
        if (Source is IDisposable disposable) disposable.Dispose();
    }
}
