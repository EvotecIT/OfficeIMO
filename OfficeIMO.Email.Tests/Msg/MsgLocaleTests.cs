using OfficeIMO.Email.Store;

namespace OfficeIMO.Email.Tests;

public sealed class MsgLocaleTests {
    [Theory]
    [InlineData(0, "en-US")]
    [InlineData(null, "en-US")]
    [InlineData(1033, "en-US")]
    [InlineData(1036, "fr-FR")]
    [InlineData(1041, "ja-JP")]
    [InlineData(int.MaxValue, "en-US")]
    public void DerivesLanguageWithoutChangingTheStoredLocale(int? locale, string language) {
        EmailDocument source = CreateMessage("Locale", locale);

        byte[] bytes = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);
        EmailReadResult result = new EmailDocumentReader().Read(bytes);

        Assert.Equal(language, result.Document.Mapi.GetValueOrDefault(MapiKnownProperties.PidName.AcceptLanguage));
        AssertMessage(result.Document, "Locale", locale ?? 1033);
        Assert.Equal(locale, source.MessageMetadata.LocaleId);
        Assert.DoesNotContain(result.Diagnostics, item => item.Severity == EmailDiagnosticSeverity.Error);
        if (locale.HasValue) {
            Assert.Equal(locale.Value, source.Mapi.GetValueOrDefault(MapiKnownProperties.PidTag.MessageLocaleId));
        } else {
            Assert.DoesNotContain(source.MapiProperties, item => item.PropertyId == 0x3FF1);
        }
    }

    [Theory]
    [InlineData("Accept-Language")]
    [InlineData("X-Accept-Language")]
    public void ExplicitLanguageHeaderTakesPrecedenceOverUnspecifiedLocale(string headerName) {
        EmailDocument source = CreateMessage("Explicit language", 0);
        source.Headers.Add(new EmailHeader(headerName, "fr-CA"));

        byte[] bytes = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);
        EmailDocument parsed = new EmailDocumentReader().Read(bytes).Document;

        Assert.Equal("fr-CA", parsed.Mapi.GetValueOrDefault(MapiKnownProperties.PidName.AcceptLanguage));
        AssertMessage(parsed, "Explicit language", 0);
        Assert.Equal(0, source.MessageMetadata.LocaleId);
    }

    [Fact]
    public void UnspecifiedLocaleSurvivesNestedMessageRoundTrip() {
        EmailDocument source = CreateNestedMessages();

        byte[] bytes = new EmailDocumentWriter().ToBytes(source, EmailFileFormat.OutlookMsg);
        EmailDocument parsed = new EmailDocumentReader().Read(bytes).Document;

        AssertNestedMessages(parsed, assertLanguage: true);
        AssertNestedMessages(source, assertLanguage: false);
    }

    [Fact]
    public void UnspecifiedLocaleAndNestedContentSurvivePstRoundTrip() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-locale-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        string path = Path.Combine(directory, "locale.pst");
        EmailDocument source = CreateNestedMessages();
        try {
            using (EmailStorePstWriter writer = EmailStorePstWriter.Create(path,
                new EmailStorePstWriterOptions(failOnDataLoss: true))) {
                writer.AddItem(writer.AddFolder("Synthetic"), source);
                EmailStorePstWriteReport report = writer.Complete();
                Assert.Equal(1, report.ItemCount);
                Assert.False(report.HasErrors);
                Assert.False(report.HasDataLoss);
            }

            using EmailStoreSession session = EmailStoreSession.Open(path,
                new EmailStoreReaderOptions(retainAttachmentContent: true));
            EmailDocument parsed = session.ReadItem(Assert.Single(session.EnumerateItems())).Document;
            AssertNestedMessages(parsed, assertLanguage: false);
            AssertNestedMessages(source, assertLanguage: false);
        } finally {
            try { if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true); }
            catch (IOException) { }
            catch (UnauthorizedAccessException) { }
        }
    }

    private static EmailDocument CreateMessage(string subject, int? locale) {
        var document = new EmailDocument { Subject = subject, MessageClass = "IPM.Note" };
        document.Body.Text = "Synthetic body for " + subject;
        document.MessageMetadata.LocaleId = locale;
        if (locale.HasValue) {
            document.MapiProperties.Add(new MapiProperty(0x3FF1, MapiPropertyType.Integer32, locale.Value));
        }
        document.MapiProperties.Add(new MapiProperty(0x65E9, MapiPropertyType.Binary, new byte[] { 4, 8, 15, 16, 23, 42 }));
        return document;
    }

    private static EmailDocument CreateNestedMessages() {
        EmailDocument outer = CreateMessage("outer", 0);
        EmailDocument child = CreateMessage("child", 0);
        EmailDocument grandchild = CreateMessage("grandchild", 0);
        grandchild.Attachments.Add(new EmailAttachment {
            FileName = "payload.bin", Content = new byte[] { 2, 7, 1, 8 }, Length = 4
        });
        child.Attachments.Add(new EmailAttachment {
            FileName = "grandchild.msg", MapiAttachMethod = 5, EmbeddedDocument = grandchild
        });
        outer.Attachments.Add(new EmailAttachment {
            FileName = "child.msg", MapiAttachMethod = 5, EmbeddedDocument = child
        });
        return outer;
    }

    private static void AssertNestedMessages(EmailDocument document, bool assertLanguage) {
        foreach (string subject in new[] { "outer", "child", "grandchild" }) {
            AssertMessage(document, subject, 0);
            if (assertLanguage) {
                Assert.Equal("en-US", document.Mapi.GetValueOrDefault(MapiKnownProperties.PidName.AcceptLanguage));
            }
            EmailAttachment attachment = Assert.Single(document.Attachments);
            if (subject == "grandchild") {
                Assert.Equal("payload.bin", attachment.FileName);
                Assert.Equal(new byte[] { 2, 7, 1, 8 }, attachment.Content);
            } else {
                Assert.Equal(5, attachment.MapiAttachMethod);
                Assert.NotNull(attachment.EmbeddedDocument);
                document = attachment.EmbeddedDocument!;
            }
        }
    }

    private static void AssertMessage(EmailDocument document, string subject, int locale) {
        Assert.Equal(subject, document.Subject);
        Assert.Equal("Synthetic body for " + subject, document.Body.Text);
        Assert.Equal(locale, document.MessageMetadata.LocaleId);
        Assert.Equal(locale, document.Mapi.GetValueOrDefault(MapiKnownProperties.PidTag.MessageLocaleId));
        Assert.Equal(new byte[] { 4, 8, 15, 16, 23, 42 },
            document.MapiProperties.Single(item => item.PropertyId == 0x65E9).Value);
    }
}
