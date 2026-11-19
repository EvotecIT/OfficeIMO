using MimeKit;
#if NET10_0_OR_GREATER
using OneMailComposition = RecursiveAI.OneMail.Composition;
#endif

namespace OfficeIMO.Email.Benchmarks.Comparisons;

internal sealed record EmailMimeBenchmarkScale(
    string Name,
    int TextCharacters,
    int HtmlCharacters,
    int AttachmentCount,
    int AttachmentBytes);

internal static class EmailMimeComparisonCorpus {
    internal static IReadOnlyList<string> ScaleNames { get; } = ["Small", "Normal"];

    internal static EmailMimeBenchmarkScale Get(string name) => name switch {
        "Small" => new EmailMimeBenchmarkScale("Small", 512, 768, 1, 4 * 1024),
        "Normal" => new EmailMimeBenchmarkScale("Normal", 16 * 1024, 24 * 1024, 4, 32 * 1024),
        _ => throw new ArgumentException("Unknown email benchmark scale: " + name, nameof(name))
    };

    internal static EmailDocument CreateOfficeDocument(EmailMimeBenchmarkScale scale) {
        var document = new EmailDocument {
            Format = EmailFileFormat.Eml,
            Subject = "OfficeIMO MIME comparison " + scale.Name,
            From = new EmailAddress("sender@example.test", "Benchmark Sender"),
            MessageId = "officeimo-mime-" + scale.Name.ToLowerInvariant() + "@example.test",
            Date = new DateTimeOffset(2026, 8, 24, 9, 30, 0, TimeSpan.Zero)
        };
        document.Recipients.Add(new EmailRecipient(
            EmailRecipientKind.To,
            new EmailAddress("reader@example.test", "Benchmark Reader")));
        document.Recipients.Add(new EmailRecipient(
            EmailRecipientKind.Cc,
            new EmailAddress("archive@example.test", "Benchmark Archive")));
        document.Body.Text = CreateText(scale.TextCharacters);
        document.Body.Html = CreateHtml(scale.HtmlCharacters);
        for (int index = 0; index < scale.AttachmentCount; index++) {
            byte[] content = CreateAttachment(index, scale.AttachmentBytes);
            document.Attachments.Add(new EmailAttachment {
                FileName = $"payload-{index + 1:D2}.bin",
                ContentType = "application/octet-stream",
                Content = content,
                Length = content.Length
            });
        }
        return document;
    }

    internal static MimeMessage CreateMimeMessage(EmailMimeBenchmarkScale scale) {
        var message = new MimeMessage {
            Subject = "OfficeIMO MIME comparison " + scale.Name,
            MessageId = "officeimo-mime-" + scale.Name.ToLowerInvariant() + "@example.test",
            Date = new DateTimeOffset(2026, 8, 24, 9, 30, 0, TimeSpan.Zero)
        };
        message.From.Add(new MailboxAddress("Benchmark Sender", "sender@example.test"));
        message.To.Add(new MailboxAddress("Benchmark Reader", "reader@example.test"));
        message.Cc.Add(new MailboxAddress("Benchmark Archive", "archive@example.test"));
        var body = new BodyBuilder {
            TextBody = CreateText(scale.TextCharacters),
            HtmlBody = CreateHtml(scale.HtmlCharacters)
        };
        for (int index = 0; index < scale.AttachmentCount; index++) {
            body.Attachments.Add(
                $"payload-{index + 1:D2}.bin",
                CreateAttachment(index, scale.AttachmentBytes),
                ContentType.Parse("application/octet-stream"));
        }
        message.Body = body.ToMessageBody();
        return message;
    }

#if NET10_0_OR_GREATER
    internal static OneMailComposition.OutgoingMessage CreateOneMailMessage(EmailMimeBenchmarkScale scale) {
        var attachments = new List<OneMailComposition.OutgoingAttachment>(scale.AttachmentCount);
        for (int index = 0; index < scale.AttachmentCount; index++) {
            attachments.Add(new OneMailComposition.OutgoingAttachment(
                $"payload-{index + 1:D2}.bin",
                CreateAttachment(index, scale.AttachmentBytes),
                "application/octet-stream"));
        }
        return new OneMailComposition.OutgoingMessage {
            From = "Benchmark Sender <sender@example.test>",
            To = ["Benchmark Reader <reader@example.test>"],
            Cc = ["Benchmark Archive <archive@example.test>"],
            Subject = "OfficeIMO MIME comparison " + scale.Name,
            MessageId = "officeimo-mime-" + scale.Name.ToLowerInvariant() + "@example.test",
            Date = new DateTimeOffset(2026, 8, 24, 9, 30, 0, TimeSpan.Zero),
            Text = CreateText(scale.TextCharacters),
            Html = CreateHtml(scale.HtmlCharacters),
            Attachments = attachments
        };
    }

    /// <summary>Builds the MIME tree from the prepared message and serializes it, as the other lanes do from their models.</summary>
    internal static byte[] WriteOneMail(OneMailComposition.OutgoingMessage message) => message.ToEntity().ToArray();

#endif
    internal static byte[] WriteMimeKit(MimeMessage message) {
        using var output = new MemoryStream();
        message.WriteTo(output);
        return output.ToArray();
    }

    private static string CreateText(int characters) =>
        CreateRepeated("Deterministic MIME text body with Unicode zażółć and line content.\n", characters);

    private static string CreateHtml(int characters) =>
        "<html><body><p>" +
        CreateRepeated("Deterministic <strong>MIME</strong> HTML body with Unicode zażółć. ", characters) +
        "</p></body></html>";

    private static string CreateRepeated(string seed, int characters) {
        var builder = new System.Text.StringBuilder(characters + seed.Length);
        while (builder.Length < characters) builder.Append(seed);
        if (builder.Length > characters) builder.Length = characters;
        return builder.ToString();
    }

    private static byte[] CreateAttachment(int index, int length) {
        var content = new byte[length];
        for (int offset = 0; offset < content.Length; offset++) {
            content[offset] = (byte)((offset * 31 + index * 17 + 11) & 0xFF);
        }
        return content;
    }
}
