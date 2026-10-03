using OfficeIMO.Email.Store;

namespace OfficeIMO.Email.Tests;

public sealed class EmbeddedSemanticLossTests {
    public static IEnumerable<object[]> LossCases() {
        foreach (EmailConversionLossPolicy policy in new[] { EmailConversionLossPolicy.Block, EmailConversionLossPolicy.Warn, EmailConversionLossPolicy.Allow })
            foreach (string kind in new[] { "Appointment", "Task", "Contact", "Journal", "Protected", "Incomplete" })
                yield return new object[] { policy, kind };
    }

    [Theory]
    [MemberData(nameof(LossCases))]
    public async Task EmbeddedSemanticAndProtectionLossHonorPolicyBeforeWriting(EmailConversionLossPolicy policy, string kind) {
        EmailDocument child = Child(kind);
        var parent = new EmailDocument { Subject = "Parent" };
        parent.Attachments.Add(new EmailAttachment { EmbeddedDocument = child, FileName = "child.eml" });
        var root = new EmailDocument { Subject = "Root" };
        root.Attachments.Add(new EmailAttachment { EmbeddedDocument = parent, FileName = "parent.eml" });
        EmailFileFormat[] formats = kind == "Protected" || kind == "Incomplete"
            ? new[] { EmailFileFormat.Eml, EmailFileFormat.Emlx, EmailFileFormat.OutlookMsg, EmailFileFormat.OutlookTemplate, EmailFileFormat.Tnef }
            : new[] { EmailFileFormat.Eml, EmailFileFormat.Emlx };
        foreach (EmailFileFormat format in formats) {
            if (kind == "Incomplete" && (format == EmailFileFormat.Eml || format == EmailFileFormat.Emlx)) continue;
            var options = new EmailWriterOptions(conversionLossPolicy: policy);
            bool blocked = policy == EmailConversionLossPolicy.Block;
            string code = Code(kind);
            foreach (bool asynchronous in new[] { false, true }) {
                using var output = new MemoryStream();
                output.WriteByte(123);
                EmailWriteResult result;
                if (format == EmailFileFormat.Emlx) {
                    var writer = new EmailStoreEmlxWriter(new EmailStoreEmlxWriterOptions(options));
                    result = asynchronous ? await writer.WriteAsync(root, output) : writer.Write(root, output);
                } else {
                    var writer = new EmailDocumentWriter(options);
                    Assert.Equal(!blocked, writer.AnalyzeConversion(root, format).CanWrite);
                    result = asynchronous ? await writer.WriteAsync(root, output, format) : writer.Write(root, output, format);
                }
                Assert.Equal(blocked ? EmailConversionLossDisposition.Blocked : EmailConversionLossDisposition.Accepted, result.LossDisposition);
                Assert.Contains(result.Diagnostics, item => item.Code == code && item.Location?.StartsWith("attachment/0/attachment/0/", StringComparison.Ordinal) == true);
                if (blocked) Assert.Equal(new byte[] { 123 }, output.ToArray());
                else Assert.True(output.Length > 1);
            }
        }
    }

    [Fact]
    public void ExactRootSourceReuseAlignsAnalysisAndWritingDespiteIncompleteProjection() {
        byte[] source = IncompleteCalendar();
        EmailDocument document = new EmailDocumentReader(new EmailReaderOptions(preserveRawSource: true)).Read(source).Document;
        Assert.True(document.MimeSemanticProjectionIsIncomplete);
        var writer = new EmailDocumentWriter(new EmailWriterOptions(usePreservedRawSource: true));
        Assert.True(writer.AnalyzeConversion(document, EmailFileFormat.Eml).CanWrite);
        Assert.False(writer.AnalyzeConversion(document, EmailFileFormat.Eml).HasPotentialDataLoss);
        byte[] bytes = writer.ToBytes(document, EmailFileFormat.Eml, out EmailWriteResult result);
        result.RequireNoLoss();
        Assert.Equal(source, bytes);
    }

    private static EmailDocument Child(string kind) {
        if (kind == "Protected") return new EmailDocumentReader(new EmailReaderOptions(preserveRawSource: true)).Read(Encoding.ASCII.GetBytes(
            "Subject: Signed\r\nMIME-Version: 1.0\r\nContent-Type: multipart/signed; protocol=\"application/pkcs7-signature\"; boundary=\"s\"\r\n\r\n" +
            "--s\r\nContent-Type: text/plain\r\n\r\nbody\r\n--s\r\nContent-Type: application/pkcs7-signature\r\n\r\nsignature\r\n--s--\r\n")).Document;
        if (kind == "Incomplete") {
            EmailDocument document = new EmailDocumentReader().Read(IncompleteCalendar()).Document;
            Assert.True(document.MimeSemanticProjectionIsIncomplete);
            return document;
        }
        var child = new EmailDocument { Subject = kind };
        if (kind == "Appointment") { child.OutlookItemKind = OutlookItemKind.Appointment; child.Appointment = new OutlookAppointment(); }
        else if (kind == "Task") { child.OutlookItemKind = OutlookItemKind.Task; child.Task = new OutlookTask { IsRecurring = true }; }
        else if (kind == "Contact") child.OutlookItemKind = OutlookItemKind.Contact;
        else child.OutlookItemKind = OutlookItemKind.Journal;
        return child;
    }

    private static string Code(string kind) => kind == "Appointment" ? "EMAIL_ICALENDAR_START_REQUIRED" :
        kind == "Task" ? "EMAIL_ICALENDAR_OPAQUE_TASK_RECURRENCE" : kind == "Contact" ? "EMAIL_VCARD_OPAQUE_CONTACT_IDENTITY" :
        kind == "Protected" ? "EMAIL_PROTECTED_CONTENT_REWRITE" : kind == "Incomplete" ? "EMAIL_STORE_SEMANTIC_PROJECTION_INCOMPLETE" :
        "EMAIL_OUTLOOK_ITEM_EML_REPRESENTATION_MISSING";

    private static byte[] IncompleteCalendar() => Encoding.ASCII.GetBytes(
        "Content-Type: text/calendar; charset=utf-8\r\n\r\nBEGIN:VCALENDAR\r\nVERSION:2.0\r\n" +
        "BEGIN:VEVENT\r\nUID:range@example.com\r\nDTSTART:20260701T090000Z\r\nDTEND:20260701T100000Z\r\nRRULE:FREQ=DAILY;COUNT=5\r\nSUMMARY:Series\r\nEND:VEVENT\r\n" +
        "BEGIN:VEVENT\r\nUID:range@example.com\r\nRECURRENCE-ID;RANGE=THISANDFUTURE:20260703T090000Z\r\n" +
        "DTSTART:20260703T110000Z\r\nDTEND:20260703T120000Z\r\nSUMMARY:Moved future series\r\nEND:VEVENT\r\nEND:VCALENDAR\r\n");
}
