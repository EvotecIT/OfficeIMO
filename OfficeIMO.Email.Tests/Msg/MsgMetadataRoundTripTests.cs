using OfficeIMO.Email;
using Xunit;

namespace OfficeIMO.Email.Tests;

public sealed class MsgMetadataRoundTripTests {
    [Fact]
    public void RoundTripsMessageRecipientAndAttachmentMetadata() {
        DateTimeOffset created = new DateTimeOffset(2026, 7, 11, 8, 15, 0, TimeSpan.Zero);
        DateTimeOffset modified = created.AddMinutes(5);
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            MessageClass = "IPM.Note",
            Subject = "RE: Metadata",
            From = new EmailAddress("owner@example.com", "Owner") { AddressType = "SMTP" },
            ReceivedBy = new EmailAddress("reader@example.com", "Reader") { AddressType = "SMTP" },
            ReceivedRepresenting = new EmailAddress("team@example.com", "Team") { AddressType = "SMTP" },
            OutlookCodePage = 1250
        };
        source.MessageMetadata.SubjectPrefix = "RE: ";
        source.MessageMetadata.NormalizedSubject = "Metadata";
        source.MessageMetadata.ConversationTopic = "Metadata thread";
        source.MessageMetadata.ConversationIndex = new byte[] { 1, 2, 3, 4 };
        source.MessageMetadata.InternetReferences = "<parent@example.com>";
        source.MessageMetadata.InReplyToId = "<parent@example.com>";
        source.MessageMetadata.Importance = EmailMessageImportance.High;
        source.MessageMetadata.Priority = EmailMessagePriority.Urgent;
        source.MessageMetadata.IconIndex = 0x00000105;
        source.MessageMetadata.IsDraft = true;
        source.MessageMetadata.IsRead = true;
        source.MessageMetadata.ReadReceiptRequested = true;
        source.MessageMetadata.DeliveryReceiptRequested = true;
        source.MessageMetadata.Sensitivity = 2;
        source.MessageMetadata.OriginalSensitivity = 1;
        source.MessageMetadata.LastModifierName = "Modifier";
        source.MessageMetadata.LocaleId = 1045;
        source.MessageMetadata.ConversationId = new byte[] { 9, 8, 7 };
        source.MessageMetadata.EditorFormat = 2;
        source.MessageMetadata.ReactionsSummary = new byte[] { 0x76, 0x01, 0x00 };
        source.MessageMetadata.OwnerReactionHistory = Encoding.UTF8.GetBytes("[]");
        source.MessageMetadata.OwnerReactionType = "like";
        source.MessageMetadata.OwnerReactionTime = created;
        source.MessageMetadata.ReactionsCount = 1;
        source.MessageMetadata.CreatedDate = created;
        source.MessageMetadata.ModifiedDate = modified;
        source.MessageMetadata.Categories.Add("Customer");
        source.MessageMetadata.Categories.Add("Follow up");

        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.To,
            new EmailAddress("person@example.com", "Person")));
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.Resource,
            new EmailAddress("projector@example.com", "Projector")));
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.Room,
            new EmailAddress("boardroom@example.com", "Board room")));
        source.Recipients.Add(new EmailRecipient(EmailRecipientKind.ReplyTo,
            new EmailAddress("replies@example.com", "Replies")));
        source.Attachments.Add(new EmailAttachment {
            FileName = "inline.png",
            ContentType = "image/png",
            ContentId = "logo",
            ContentLocation = "images/logo.png",
            Content = new byte[] { 1, 2, 3 },
            Length = 3,
            IsInline = true,
            IsHidden = true,
            IsContactPhoto = true,
            RenderingPosition = 12,
            CreatedDate = created,
            ModifiedDate = modified,
            LinkedPath = "images/logo.png"
        });

        byte[] bytes = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);
        EmailReadResult result = new EmailDocumentReader().Read(bytes);
        EmailDocument parsed = result.Document;

        Assert.Equal("RE: ", parsed.MessageMetadata.SubjectPrefix);
        Assert.Equal("Metadata", parsed.MessageMetadata.NormalizedSubject);
        Assert.Equal("Metadata thread", parsed.MessageMetadata.ConversationTopic);
        Assert.Equal(new byte[] { 1, 2, 3, 4 }, parsed.MessageMetadata.ConversationIndex);
        Assert.Equal("<parent@example.com>", parsed.MessageMetadata.InternetReferences);
        Assert.Equal("<parent@example.com>", parsed.MessageMetadata.InReplyToId);
        Assert.Equal(EmailMessageImportance.High, parsed.MessageMetadata.Importance);
        Assert.Equal(EmailMessagePriority.Urgent, parsed.MessageMetadata.Priority);
        Assert.Equal(0x00000105, parsed.MessageMetadata.IconIndex);
        Assert.True(parsed.MessageMetadata.IsDraft);
        Assert.True(parsed.MessageMetadata.IsRead);
        Assert.True(parsed.MessageMetadata.ReadReceiptRequested);
        Assert.True(parsed.MessageMetadata.DeliveryReceiptRequested);
        Assert.Equal(2, parsed.MessageMetadata.Sensitivity);
        Assert.Equal(1, parsed.MessageMetadata.OriginalSensitivity);
        Assert.Equal("Modifier", parsed.MessageMetadata.LastModifierName);
        Assert.Equal(1045, parsed.MessageMetadata.LocaleId);
        Assert.Equal(new byte[] { 9, 8, 7 }, parsed.MessageMetadata.ConversationId);
        Assert.Equal(2, parsed.MessageMetadata.EditorFormat);
        Assert.Equal(new byte[] { 0x76, 0x01, 0x00 }, parsed.MessageMetadata.ReactionsSummary);
        Assert.Equal("[]", Encoding.UTF8.GetString(parsed.MessageMetadata.OwnerReactionHistory!));
        Assert.Equal("like", parsed.MessageMetadata.OwnerReactionType);
        Assert.Equal(created, parsed.MessageMetadata.OwnerReactionTime);
        Assert.Equal(1, parsed.MessageMetadata.ReactionsCount);
        Assert.Equal(created, parsed.MessageMetadata.CreatedDate);
        Assert.Equal(modified, parsed.MessageMetadata.ModifiedDate);
        Assert.Equal(new[] { "Customer", "Follow up" }, parsed.MessageMetadata.Categories);

        Assert.Contains(parsed.Recipients, recipient => recipient.Kind == EmailRecipientKind.Resource &&
            recipient.Address.Address == "projector@example.com");
        Assert.Contains(parsed.Recipients, recipient => recipient.Kind == EmailRecipientKind.Room &&
            recipient.Address.Address == "boardroom@example.com");
        Assert.Contains(parsed.Recipients, recipient => recipient.Kind == EmailRecipientKind.ReplyTo &&
            recipient.Address.Address == "replies@example.com");
        Assert.Equal("reader@example.com", parsed.ReceivedBy!.Address);
        Assert.Equal("team@example.com", parsed.ReceivedRepresenting!.Address);

        EmailAttachment attachment = Assert.Single(parsed.Attachments);
        Assert.True(attachment.IsInline);
        Assert.True(attachment.IsHidden);
        Assert.True(attachment.IsContactPhoto);
        Assert.Equal(12, attachment.RenderingPosition);
        Assert.Equal(created, attachment.CreatedDate);
        Assert.Equal(modified, attachment.ModifiedDate);
        Assert.Equal("images/logo.png", attachment.LinkedPath);

        using MemoryStream stream = new MemoryStream(bytes);
        using var oracle = new global::MsgReader.Outlook.Storage.Message(stream, FileAccess.Read, true);
        Assert.Equal(3, oracle.Recipients!.Count);
        Assert.Contains(oracle.Recipients, recipient => recipient.Type == global::MsgReader.Outlook.RecipientType.Resource);
        Assert.Contains(oracle.Recipients, recipient => recipient.Type == global::MsgReader.Outlook.RecipientType.Room);
        Assert.Equal("reader@example.com", oracle.ReceivedBy.Email);
        Assert.Equal(new[] { "Customer", "Follow up" }, oracle.Categories);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
    }

    [Fact]
    public void PreservesUntouchedString8BytesAcrossReadWrite() {
        byte[] shiftJis = new byte[] { 0x93, 0xFA, 0x96, 0x7B };
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            Subject = "String8",
            OutlookCodePage = 932
        };
        source.MapiProperties.Add(new MapiProperty(0x66AB, MapiPropertyType.String8, "日本") {
            RawData = shiftJis
        });

        byte[] first = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);
        EmailDocument parsed = new EmailDocumentReader().Read(first).Document;
        MapiProperty property = parsed.MapiProperties.Single(item => item.PropertyId == 0x66AB);

        Assert.Equal("日本", property.Value);
        Assert.Equal(shiftJis, property.RawData);

        byte[] second = new EmailDocumentWriter().ToBytes(parsed, EmailFileFormat.OutlookMsg);
        MapiProperty reparsed = new EmailDocumentReader().Read(second).Document.MapiProperties
            .Single(item => item.PropertyId == 0x66AB);
        Assert.Equal("日本", reparsed.Value);
        Assert.Equal(shiftJis, reparsed.RawData);
    }

    [Fact]
    public void EncodesNewString8AndHtmlValuesWithTheDeclaredCodePage() {
        var source = new EmailDocument {
            Format = EmailFileFormat.OutlookMsg,
            Subject = "Code page",
            OutlookCodePage = 932
        };
        source.Body.Html = "<p>日本</p>";
        source.MapiProperties.Add(new MapiProperty(0x66AB, MapiPropertyType.String8, "日本"));

        EmailReadResult result = new EmailDocumentReader().Read(
            new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg));

        Assert.Equal(932, result.Document.OutlookCodePage);
        Assert.Equal("<p>日本</p>", result.Document.Body.Html);
        Assert.Equal("日本", result.Document.MapiProperties.Single(property => property.PropertyId == 0x66AB).Value);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error);
    }
}
