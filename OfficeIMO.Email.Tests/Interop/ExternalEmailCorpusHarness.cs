using MimeKit;
using MimeKit.Tnef;
using OfficeIMO.Email;

namespace OfficeIMO.Email.Tests;

internal sealed class ExternalCorpusResult {
    private readonly List<string> _failures = new List<string>();

    internal int CandidateArtifacts { get; private set; }
    internal int ApplicableArtifacts { get; private set; }
    internal int SkippedArtifacts { get; private set; }
    internal IReadOnlyList<string> Failures => _failures;

    internal void Run(string category, string path, Func<bool> action) {
        CandidateArtifacts++;
        try {
            if (action()) ApplicableArtifacts++;
            else SkippedArtifacts++;
        } catch (Exception exception) {
            _failures.Add(string.Concat(category, ": ", Path.GetFileName(path), ": ",
                exception.GetType().Name, ": ", exception.Message, Environment.NewLine, exception.StackTrace));
        }
    }

    internal string FormatFailures() => string.Join(Environment.NewLine, _failures);
}

internal static class ExternalEmailCorpusHarness {
    private static readonly EmailDocumentReader Reader = new EmailDocumentReader();
    private static readonly EmailDocumentWriter Writer = new EmailDocumentWriter(new EmailWriterOptions(
        conversionLossPolicy: EmailConversionLossPolicy.Warn));

    internal static string? FindRepository(string name) {
        foreach (string? root in new[] {
            Environment.GetEnvironmentVariable("OFFICEIMO_EMAIL_CORPUS_ROOT"),
            Environment.GetEnvironmentVariable("EVOTEC_GITHUB_ROOT")
        }) {
            if (string.IsNullOrWhiteSpace(root)) continue;
            string candidate = Path.Combine(root!, name);
            if (Directory.Exists(candidate)) return candidate;
        }
        return null;
    }

    internal static ExternalCorpusResult RunMsgReader(string repository) {
        var result = new ExternalCorpusResult();
        string samples = Path.Combine(repository, "MsgReaderTests", "SampleFiles");
        if (!Directory.Exists(samples)) return result;

        foreach (string path in Directory.GetFiles(samples, "*.eml", SearchOption.AllDirectories)
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)) {
            result.Run("MsgReader EML", path, () => ValidateMime(path));
        }
        foreach (string path in Directory.GetFiles(samples, "*.msg", SearchOption.AllDirectories)
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase)) {
            result.Run("MsgReader MSG", path, () => ValidateOutlookMsg(path));
        }
        return result;
    }

    internal static ExternalCorpusResult RunMimeKit(string repository) {
        var result = new ExternalCorpusResult();
        string data = Path.Combine(repository, "UnitTests", "TestData");
        if (!Directory.Exists(data)) return result;

        foreach (string path in FindMimeKitMessages(data)) {
            result.Run("MimeKit MIME", path, () => ValidateMime(path));
        }
        foreach (string path in Directory.GetFiles(Path.Combine(data, "tnef"), "*.tnef",
            SearchOption.TopDirectoryOnly).OrderBy(path => path, StringComparer.OrdinalIgnoreCase)) {
            result.Run("MimeKit TNEF", path, () => ValidateTnef(path));
        }
        foreach (string path in Directory.GetFiles(Path.Combine(data, "mbox"), "*.mbox.txt",
            SearchOption.TopDirectoryOnly).OrderBy(path => path, StringComparer.OrdinalIgnoreCase)) {
            result.Run("MimeKit mbox", path, () => ValidateMbox(path));
        }
        return result;
    }

    private static IEnumerable<string> FindMimeKitMessages(string data) {
        var paths = new List<string>();
        AddFiles(paths, Path.Combine(data, "messages"), "*.eml");
        AddFiles(paths, Path.Combine(data, "messages"), "*.txt", path =>
            !Path.GetFileName(path).StartsWith("body.", StringComparison.OrdinalIgnoreCase));
        AddFiles(paths, Path.Combine(data, "dkim"), "*.msg");
        AddFiles(paths, Path.Combine(data, "yenc"), "*.msg");
        AddFiles(paths, Path.Combine(data, "openpgp"), "*.eml");
        AddFiles(paths, Path.Combine(data, "partial"), "*.eml");
        AddFiles(paths, Path.Combine(data, "tnef"), "*.eml");
        return paths.Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(path => path, StringComparer.OrdinalIgnoreCase);
    }

    private static void AddFiles(ICollection<string> paths, string directory, string pattern,
        Func<string, bool>? predicate = null) {
        if (!Directory.Exists(directory)) return;
        foreach (string path in Directory.GetFiles(directory, pattern, SearchOption.TopDirectoryOnly)) {
            if (predicate == null || predicate(path)) paths.Add(path);
        }
    }

    private static bool ValidateMime(string path) {
        using (FileStream formatStream = File.OpenRead(path)) {
            if (EmailDocumentReader.DetectFormat(formatStream) == EmailFileFormat.Mbox) return ValidateMbox(path);
        }
        MimeMessage oracle;
        try {
            oracle = MimeMessage.Load(path);
        } catch (FormatException) {
            return false;
        }
        using (oracle) {
            EmailReadResult read = Reader.Read(path);
            EnsureReadable(read, path);
            Check(read.Document.Format == EmailFileFormat.Eml, "OfficeIMO did not classify the artifact as EML.");
            Check(EqualText(NormalizeOracleText(oracle.Subject), read.Document.Subject), string.Concat(
                "The decoded subject differs from MimeKit (OfficeIMO: '", Escape(read.Document.Subject),
                "'; MimeKit: '", Escape(oracle.Subject), "')."));
            if (!string.IsNullOrWhiteSpace(oracle.TextBody)) {
                string? oracleText = NormalizeOracleText(oracle.TextBody);
                Check(!string.IsNullOrWhiteSpace(read.Document.Body.Text),
                    "MimeKit found a text body but OfficeIMO did not.");
                if (IsYEncBody(oracle.TextBody)) {
                    Check(read.Document.Body.Text!.Contains("=ybegin", StringComparison.Ordinal) &&
                        read.Document.Body.Text.Contains("=yend", StringComparison.Ordinal),
                        "OfficeIMO did not preserve the yEnc framing in the non-MIME Usenet body.");
                } else {
                    Check(EqualBody(oracleText, read.Document.Body.Text),
                        DescribeBodyDifference("MimeKit text", oracleText, read.Document.Body.Text));
                }
            }
            if (!string.IsNullOrWhiteSpace(oracle.HtmlBody)) {
                string? oracleHtml = NormalizeOracleText(oracle.HtmlBody);
                Check(!string.IsNullOrWhiteSpace(read.Document.Body.Html),
                    "MimeKit found an HTML body but OfficeIMO did not.");
                Check(EqualBody(oracleHtml, read.Document.Body.Html),
                    DescribeBodyDifference("MimeKit HTML", oracleHtml, read.Document.Body.Html));
            }
            int oracleAttachments = oracle.Attachments.Count();
            Check(read.Document.Attachments.Count >= oracleAttachments,
                "OfficeIMO exposed fewer attachments than MimeKit.");
            int oracleRecipients = oracle.To.Mailboxes.Count() + oracle.Cc.Mailboxes.Count() +
                oracle.Bcc.Mailboxes.Count();
            int officeRecipients = read.Document.Recipients.Count(recipient =>
                recipient.Kind == EmailRecipientKind.To || recipient.Kind == EmailRecipientKind.Cc ||
                recipient.Kind == EmailRecipientKind.Bcc);
            Check(officeRecipients == oracleRecipients,
                string.Concat("Recipient count differs from MimeKit (", officeRecipients, " vs ",
                    oracleRecipients, "). OfficeIMO: ",
                    string.Join(" | ", read.Document.Recipients.Select(recipient =>
                        string.Concat(recipient.Kind, ":", recipient.Address.Address, ":", recipient.Address.DisplayName))),
                    "; MimeKit: ", string.Join(" | ", oracle.To.Mailboxes.Concat(oracle.Cc.Mailboxes)
                    .Concat(oracle.Bcc.Mailboxes).Select(mailbox => mailbox.ToString())), "."));
            CheckAddresses("To", oracle.To.Mailboxes.Select(mailbox => mailbox.Address), read.Document,
                EmailRecipientKind.To);
            CheckAddresses("Cc", oracle.Cc.Mailboxes.Select(mailbox => mailbox.Address), read.Document,
                EmailRecipientKind.Cc);
            CheckAddresses("Bcc", oracle.Bcc.Mailboxes.Select(mailbox => mailbox.Address), read.Document,
                EmailRecipientKind.Bcc);
            MailboxAddress? oracleFrom = oracle.From.Mailboxes.FirstOrDefault();
            if (oracleFrom != null) {
                Check(string.Equals(oracleFrom.Address, read.Document.From?.Address,
                        StringComparison.OrdinalIgnoreCase),
                    string.Concat("From address differs from MimeKit ('", read.Document.From?.Address,
                        "' vs '", oracleFrom.Address, "')."));
            }

            ValidateAllOutputFormats(read.Document, path);
            return true;
        }
    }

    private static bool ValidateOutlookMsg(string path) {
        using var input = File.OpenRead(path);
        using var oracle = new global::MsgReader.Outlook.Storage.Message(input, FileAccess.Read, true);
        EmailReadResult read = Reader.Read(path);
        EnsureReadable(read, path);
        Check(read.Document.Format == EmailFileFormat.OutlookMsg,
            "OfficeIMO did not classify the artifact as Outlook MSG.");
        Check(EqualText(oracle.Subject, read.Document.Subject), "The decoded subject differs from MsgReader.");
        if (!string.IsNullOrWhiteSpace(oracle.BodyText)) {
            Check(!string.IsNullOrWhiteSpace(read.Document.Body.Text),
                "MsgReader found a text body but OfficeIMO did not.");
            Check(EqualMsgBody(oracle.BodyText, read.Document.Body.Text),
                DescribeBodyDifference("MsgReader text", NormalizeMsgBody(oracle.BodyText),
                    NormalizeMsgBody(read.Document.Body.Text)));
        }
        Check(read.Document.Attachments.Count == oracle.Attachments.Count,
            string.Concat("Attachment count differs from MsgReader (", read.Document.Attachments.Count, " vs ",
                oracle.Attachments.Count, ")."));
        Check(read.Document.Recipients.Count(recipient => recipient.Kind != EmailRecipientKind.ReplyTo) ==
            oracle.Recipients.Count,
            "Recipient count differs from MsgReader.");

        ValidateAllOutputFormats(read.Document, path);
        return true;
    }

    private static bool ValidateTnef(string path) {
        List<MimeEntity> oracleAttachments;
        try {
            using var content = File.OpenRead(path);
            var tnef = new TnefPart { Content = new MimeContent(content) };
            oracleAttachments = tnef.ExtractAttachments().ToList();
        } catch (FormatException) {
            return false;
        }

        EmailReadResult read = Reader.Read(path);
        EnsureReadable(read, path);
        Check(read.Document.Format == EmailFileFormat.Tnef, "OfficeIMO did not classify the artifact as TNEF.");
        MimePart[] oracleFiles = oracleAttachments.OfType<MimePart>()
            .Where(part => !(part is TextPart) || !string.IsNullOrWhiteSpace(part.FileName))
            .GroupBy(part => string.Concat(part.FileName, "\0", part.ContentDisposition?.Size),
                StringComparer.OrdinalIgnoreCase)
            .Select(group => group.First())
            .ToArray();
        EmailAttachment[] officeAttachments = read.Document.Attachments.ToArray();
        Check(officeAttachments.Length >= oracleFiles.Length,
            string.Concat("OfficeIMO exposed fewer TNEF attachments than MimeKit (",
                officeAttachments.Length, " vs ", oracleFiles.Length, "). OfficeIMO: ",
                string.Join(" | ", officeAttachments.Select(attachment => string.Concat(
                    attachment.FileName, ":", attachment.ContentType, ":", attachment.Length))),
                "; MimeKit: ", string.Join(" | ", oracleAttachments.Select(entity => string.Concat(
                    entity.ContentType.MimeType, ":", entity.ContentDisposition?.FileName))), "."));

        ValidateAllOutputFormats(read.Document, path);
        return true;
    }

    private static bool ValidateMbox(string path) {
        int oracleCount = CountMimeKitMbox(path);
        var entries = new EmailMailboxReader().ReadEntries(path).ToArray();
        Check(entries.All(entry => !entry.HasErrors), string.Concat(
            "OfficeIMO reported an error while streaming the mailbox: ",
            string.Join(" | ", entries.SelectMany((entry, index) => entry.Diagnostics
                .Where(diagnostic => diagnostic.Severity == EmailDiagnosticSeverity.Error)
                .Select(diagnostic => string.Concat("message[", index, "] ", diagnostic.Code, ": ", diagnostic.Message))))));
        Check(entries.Length == oracleCount, string.Concat("Mailbox message count differs from MimeKit (",
            entries.Length, " vs ", oracleCount, ")."));

        var mailbox = new EmailMailbox();
        foreach (EmailMailboxEntryReadResult entry in entries) mailbox.Messages.Add(entry.Entry);
        byte[] rewritten = new EmailMailboxWriter().ToBytes(mailbox);
        using var stream = new MemoryStream(rewritten);
        int rewrittenCount = CountMimeKitMbox(stream);
        Check(rewrittenCount == entries.Length, "MimeKit did not reopen every rewritten mailbox entry.");
        return true;
    }

    private static void ValidateAllOutputFormats(EmailDocument document, string sourcePath) {
        byte[] eml = Writer.ToBytes(document, EmailFileFormat.Eml);
        using (var stream = new MemoryStream(eml)) {
            using MimeMessage oracle = MimeMessage.Load(stream);
            Check(EqualText(document.Subject, oracle.Subject), "MimeKit read a different rewritten EML subject.");
        }
        EnsureReadable(Reader.Read(eml), string.Concat(sourcePath, " -> EML"));

        byte[] msg = Writer.ToBytes(document, EmailFileFormat.OutlookMsg);
        EnsureReadable(Reader.Read(msg), string.Concat(sourcePath, " -> MSG"));
        using (var oracle = new global::MsgReader.Outlook.Storage.Message(
                   new MemoryStream(msg), FileAccess.Read, true)) {
            Check(EqualText(document.Subject, oracle.Subject), "MsgReader read a different rewritten MSG subject.");
        }

        byte[] tnef = Writer.ToBytes(document, EmailFileFormat.Tnef);
        EnsureReadable(Reader.Read(tnef), string.Concat(sourcePath, " -> TNEF"));
    }

    private static int CountMimeKitMbox(string path) {
        using var stream = File.OpenRead(path);
        return CountMimeKitMbox(stream);
    }

    private static int CountMimeKitMbox(Stream stream) {
        var parser = new global::MimeKit.MimeParser(stream, MimeFormat.Mbox);
        int count = 0;
        while (!parser.IsEndOfStream) {
            using MimeMessage message = parser.ParseMessage();
            count++;
        }
        return count;
    }

    private static void EnsureReadable(EmailReadResult result, string location) {
        EmailDiagnostic? error = result.Diagnostics.FirstOrDefault(diagnostic =>
            diagnostic.Severity == EmailDiagnosticSeverity.Error);
        if (error != null) {
            throw new InvalidDataException(string.Concat(location, ": ", error.Code, ": ", error.Message));
        }
    }

    private static bool EqualText(string? left, string? right) =>
        string.Equals(left ?? string.Empty, right ?? string.Empty, StringComparison.Ordinal);

    private static bool EqualBody(string? left, string? right) => string.Equals(
        NormalizeBody(left), NormalizeBody(right), StringComparison.Ordinal);

    private static bool EqualMsgBody(string? left, string? right) => string.Equals(
        NormalizeMsgBody(left), NormalizeMsgBody(right), StringComparison.Ordinal);

    private static string NormalizeBody(string? value) => (value ?? string.Empty)
        .Replace("\r\n", "\n").Replace("\r", "\n").TrimEnd('\n');

    private static string NormalizeMsgBody(string? value) {
        string[] lines = NormalizeBody(value).Split('\n');
        for (int index = 0; index < lines.Length; index++) lines[index] = lines[index].TrimEnd(' ', '\t');
        return string.Join("\n", lines).TrimEnd('\n');
    }

    private static string DescribeBodyDifference(string oracleName, string? oracle, string? office) {
        string expected = NormalizeBody(oracle);
        string actual = NormalizeBody(office);
        int shared = Math.Min(expected.Length, actual.Length);
        int offset = 0;
        while (offset < shared && expected[offset] == actual[offset]) offset++;
        int start = Math.Max(0, offset - 20);
        return string.Concat("The normalized body differs from ", oracleName,
            " at character ", offset, " (OfficeIMO length ", actual.Length,
            ", oracle length ", expected.Length, "). OfficeIMO: '",
            Escape(actual.Substring(start, Math.Min(100, actual.Length - start))),
            "'; oracle: '", Escape(expected.Substring(start, Math.Min(100, expected.Length - start))), "'.");
    }

    private static void CheckAddresses(string label, IEnumerable<string> oracleAddresses,
        EmailDocument document, EmailRecipientKind kind) {
        string[] expected = oracleAddresses.Where(address => !string.IsNullOrWhiteSpace(address))
            .OrderBy(address => address, StringComparer.OrdinalIgnoreCase).ToArray();
        string[] actual = document.Recipients.Where(recipient => recipient.Kind == kind)
            .Select(recipient => recipient.Address.Address)
            .Where(address => !string.IsNullOrWhiteSpace(address)).Select(address => address!)
            .OrderBy(address => address, StringComparer.OrdinalIgnoreCase).ToArray();
        Check(expected.SequenceEqual(actual, StringComparer.OrdinalIgnoreCase), string.Concat(
            label, " addresses differ (OfficeIMO: ", string.Join(", ", actual),
            "; oracle: ", string.Join(", ", expected), ")."));
    }

    private static string? NormalizeOracleText(string? value) {
        if (string.IsNullOrEmpty(value) || value!.IndexOf('\u001b') < 0) return value;
        try {
            return Encoding.GetEncoding("iso-2022-jp").GetString(Encoding.ASCII.GetBytes(value));
        } catch (ArgumentException) {
            return value;
        }
    }

    private static bool IsYEncBody(string? value) => value != null &&
        value.IndexOf("=ybegin", StringComparison.Ordinal) >= 0 &&
        value.IndexOf("=yend", StringComparison.Ordinal) >= 0;

    private static string Escape(string? value) => (value ?? string.Empty)
        .Replace("\r", "\\r").Replace("\n", "\\n").Replace("\0", "\\0");

    private static void Check(bool condition, string message) {
        if (!condition) throw new InvalidDataException(message);
    }
}
