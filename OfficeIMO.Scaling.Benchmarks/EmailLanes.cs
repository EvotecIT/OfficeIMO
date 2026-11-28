using OfficeIMO.Email;

namespace OfficeIMO.Scaling.Benchmarks;

internal static class EmailFixtures {
    private static readonly byte[] AttachmentContent = Enumerable.Range(0, 2048).Select(value => (byte)value).ToArray();

    public static EmailDocument WithAttachments(int count) {
        var message = new EmailDocument {
            Subject = "Scaling",
            From = new EmailAddress("sender@example.com", "Sender")
        };
        message.Body.Text = "Body text.";
        message.Recipients.Add(new EmailRecipient(EmailRecipientKind.To, new EmailAddress("to@example.com", "To")));
        for (int i = 0; i < count; i++) {
            message.Attachments.Add(new EmailAttachment {
                FileName = $"file{i}.bin",
                ContentType = "application/octet-stream",
                Content = AttachmentContent
            });
        }

        return message;
    }
}

internal sealed class MsgAttachmentWriteLane : ScalingLane {
    private EmailDocument _message = null!;
    public override string Name => "Msg.Write";
    public override string Unit => "attachment";
    protected override void Prepare(int units) => _message = EmailFixtures.WithAttachments(units);
    public override object? Run() => _message.ToBytes(EmailFileFormat.OutlookMsg);

    public override void Validate(object? result) =>
        Require(EmailDocument.Load((byte[])result!).Attachments.Count == Units, "attachment count");
}

internal abstract class EmailAttachmentReadLane : ScalingLane {
    private byte[] _bytes = [];
    protected abstract EmailFileFormat Format { get; }
    public override string Unit => "attachment";
    protected override void Prepare(int units) => _bytes = EmailFixtures.WithAttachments(units).ToBytes(Format);
    public override object? Run() => EmailDocument.Load(_bytes).Attachments.Count;
    public override void Validate(object? result) => Require((int)result! == Units, "attachment count");
}

internal sealed class MsgAttachmentReadLane : EmailAttachmentReadLane {
    public override string Name => "Msg.Read";
    protected override EmailFileFormat Format => EmailFileFormat.OutlookMsg;
}

internal sealed class EmlAttachmentReadLane : EmailAttachmentReadLane {
    public override string Name => "Eml.Read";
    protected override EmailFileFormat Format => EmailFileFormat.Eml;
}

internal sealed class TnefAttachmentReadLane : EmailAttachmentReadLane {
    public override string Name => "Tnef.Read";
    protected override EmailFileFormat Format => EmailFileFormat.Tnef;
}

internal sealed class MsgRecipientWriteLane : ScalingLane {
    private EmailDocument _message = null!;
    public override string Name => "Msg.WriteRecipients";
    public override string Unit => "recipient";

    protected override void Prepare(int units) {
        _message = EmailFixtures.WithAttachments(0);
        for (int i = 1; i < units; i++) {
            _message.Recipients.Add(new EmailRecipient(EmailRecipientKind.Cc, new EmailAddress($"r{i}@example.com", $"Recipient {i}")));
        }
    }

    public override object? Run() => _message.ToBytes(EmailFileFormat.OutlookMsg);

    public override void Validate(object? result) =>
        Require(EmailDocument.Load((byte[])result!).Recipients.Count == Units, "recipient count");
}
