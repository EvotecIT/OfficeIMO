using OfficeIMO.Email;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace OfficeIMO.Reader;

internal static partial class DocumentReaderEngine {
    private static IEnumerable<ReaderChunk> ReadEmail(string path, ReaderOptions opt, CancellationToken cancellationToken) {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        EmailExtraction extraction = ExtractEmail(stream, path, opt, cancellationToken);
        return extraction.Chunks;
    }

    private static IEnumerable<ReaderChunk> ReadEmail(Stream stream, string? sourceName, ReaderOptions opt, CancellationToken cancellationToken) {
        EmailExtraction extraction = ExtractEmail(stream, NormalizeLogicalSourceName(sourceName, "message.eml"), opt, cancellationToken);
        return extraction.Chunks;
    }

    private static OfficeDocumentReadResult ReadEmailDocument(string path, ReaderOptions opt, CancellationToken cancellationToken) {
        EnforceFileSize(path, GetEffectiveEmailMaxInputBytes(opt));
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
        EmailExtraction extraction = ExtractEmail(stream, path, opt, cancellationToken);
        SourceInfo source = BuildSourceInfoFromPath(path, opt.ComputeHashes,
            cancellationToken);
        EnrichEmailChunks(extraction.Chunks, source, opt.ComputeHashes);
        return BuildEmailDocumentResult(extraction, path, BuildPathDocumentSource(path, extraction.Chunks));
    }

    private static OfficeDocumentReadResult ReadEmailDocument(Stream stream, string sourceName, ReaderOptions opt, CancellationToken cancellationToken) {
        using MemoryStream snapshot = CopyEmailToMemory(stream, opt, cancellationToken);
        return ReadEmailDocument(snapshot, sourceName, opt, cancellationToken);
    }

    private static OfficeDocumentReadResult ReadEmailDocument(MemoryStream snapshot, string sourceName, ReaderOptions opt, CancellationToken cancellationToken) {
        snapshot.Position = 0;
        EmailExtraction extraction = ExtractEmail(snapshot.ToArray(), sourceName, opt, cancellationToken);
        snapshot.Position = 0;
        SourceInfo source = BuildSourceInfoFromStream(snapshot, sourceName,
            opt.ComputeHashes, cancellationToken);
        EnrichEmailChunks(extraction.Chunks, source, opt.ComputeHashes);
        return BuildEmailDocumentResult(extraction, sourceName, BuildStreamDocumentSource(snapshot, sourceName, extraction.Chunks));
    }

    private static void EnrichEmailChunks(IList<ReaderChunk> chunks, SourceInfo source, bool computeHashes) {
        for (int index = 0; index < chunks.Count; index++) {
            chunks[index] = EnrichChunk(chunks[index], source, computeHashes);
        }
    }

    private static EmailExtraction ExtractEmail(Stream stream, string sourceName, ReaderOptions opt, CancellationToken cancellationToken) {
        using MemoryStream snapshot = CopyEmailToMemory(stream, opt, cancellationToken);
        return ExtractEmail(snapshot.ToArray(), sourceName, opt, cancellationToken);
    }

    private static EmailExtraction ExtractEmail(byte[] bytes, string sourceName, ReaderOptions opt, CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        EmailReaderOptions emailOptions = CreateEmailReaderOptions(opt);
        EmailFileFormat format = EmailDocumentReader.DetectFormat(bytes);
        var extraction = new EmailExtraction(format, sourceName);

        if (format == EmailFileFormat.Mbox) {
            var mailboxReader = new EmailMailboxReader(new EmailMailboxReaderOptions(emailOptions));
            EmailMailboxReadResult mailboxResult = mailboxReader.Read(bytes, cancellationToken);
            extraction.Diagnostics.AddRange(mailboxResult.Diagnostics);
            for (int index = 0; index < mailboxResult.Mailbox.Messages.Count; index++) {
                EmailMailboxEntry entry = mailboxResult.Mailbox.Messages[index];
                extraction.Documents.Add(entry.Document);
                extraction.MailboxEntries.Add(entry);
            }
        } else {
            EmailReadResult result = new EmailDocumentReader(emailOptions).Read(bytes, sourceName,
                cancellationToken);
            extraction.Format = result.Document.Format;
            extraction.Diagnostics.AddRange(result.Diagnostics);
            extraction.Documents.Add(result.Document);
            extraction.MailboxEntries.Add(null);
        }

        BuildEmailChunks(extraction, opt, cancellationToken);
        return extraction;
    }

    private static EmailReaderOptions CreateEmailReaderOptions(ReaderOptions opt) {
        EmailReaderOptions defaults = EmailReaderOptions.Default;
        long maxInputBytes = GetEffectiveEmailMaxInputBytes(opt);

        return new EmailReaderOptions(
            maxInputBytes: maxInputBytes,
            maxHeaderBytes: defaults.MaxHeaderBytes,
            maxHeaderCount: defaults.MaxHeaderCount,
            maxPartCount: defaults.MaxPartCount,
            maxMimeDepth: defaults.MaxMimeDepth,
            maxAttachmentBytes: Math.Min(defaults.MaxAttachmentBytes, maxInputBytes),
            maxTotalAttachmentBytes: Math.Min(defaults.MaxTotalAttachmentBytes, maxInputBytes),
            maxNestedMessageDepth: defaults.MaxNestedMessageDepth,
            includeAttachmentContent: true,
            preserveRawSource: false,
            maxCompoundDirectoryEntries: defaults.MaxCompoundDirectoryEntries,
            maxMapiPropertyCount: defaults.MaxMapiPropertyCount,
            maxDecodedPropertyBytes: Math.Min(defaults.MaxDecodedPropertyBytes, maxInputBytes),
            maxTnefAttributeCount: defaults.MaxTnefAttributeCount);
    }

    private static long GetEffectiveEmailMaxInputBytes(ReaderOptions opt) {
        long maxInputBytes = opt.MaxInputBytes.GetValueOrDefault(EmailReaderOptions.Default.MaxInputBytes);
        return maxInputBytes > 0 ? maxInputBytes : EmailReaderOptions.Default.MaxInputBytes;
    }

    private static MemoryStream CopyEmailToMemory(Stream stream, ReaderOptions opt,
        CancellationToken cancellationToken) {
        long maxInputBytes = GetEffectiveEmailMaxInputBytes(opt);
        ReaderInputLimits.EnforceSeekableStreamSize(stream, maxInputBytes);
        return CopyToMemory(stream, cancellationToken, maxInputBytes);
    }

    private static bool IsEmailArtifact(string path, ReaderOptions opt, CancellationToken cancellationToken) {
        try {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            return IsEmailArtifact(stream, opt, cancellationToken);
        } catch (OperationCanceledException) {
            throw;
        } catch (IOException) {
            return false;
        } catch (UnauthorizedAccessException) {
            return false;
        }
    }

    private static bool IsEmailArtifact(Stream stream, ReaderOptions opt, CancellationToken cancellationToken) {
        if (!stream.CanSeek) {
            return false;
        }

        long position = stream.Position;
        try {
            using MemoryStream snapshot = CopyEmailToMemory(stream, opt, cancellationToken);
            byte[] bytes = snapshot.ToArray();
            EmailFileFormat format = EmailDocumentReader.DetectFormat(bytes);
            return format != EmailFileFormat.Unknown &&
                (format != EmailFileFormat.Eml || HasDistinctiveEmailHeader(bytes));
        } catch (OperationCanceledException) {
            throw;
        } catch (IOException) {
            return false;
        } finally {
            stream.Position = position;
        }
    }

    private static bool HasDistinctiveEmailHeader(byte[] bytes) {
        int count = Math.Min(bytes.Length, 64 * 1024);
        string headerText = Encoding.ASCII.GetString(bytes, 0, count);
        string[] lines = headerText.Replace("\r\n", "\n").Split('\n');
        for (int index = 0; index < lines.Length; index++) {
            string line = lines[index];
            if (line.Length == 0) {
                break;
            }

            int colon = line.IndexOf(':');
            if (colon <= 0) {
                continue;
            }

            string name = line.Substring(0, colon).Trim();
            if (string.Equals(name, "From", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "Sender", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "To", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "Cc", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "Bcc", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "Date", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "Message-ID", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, "MIME-Version", StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }
        return false;
    }

    private sealed class EmailExtraction {
        internal EmailExtraction(EmailFileFormat format, string sourceName) {
            Format = format;
            SourceName = sourceName;
        }

        internal EmailFileFormat Format { get; set; }
        internal string SourceName { get; }
        internal List<EmailDocument> Documents { get; } = new List<EmailDocument>();
        internal List<EmailMailboxEntry?> MailboxEntries { get; } = new List<EmailMailboxEntry?>();
        internal List<string?> LogicalPaths { get; } = new List<string?>();
        internal List<EmailDiagnostic> Diagnostics { get; } = new List<EmailDiagnostic>();
        internal List<ReaderChunk> Chunks { get; } = new List<ReaderChunk>();
        internal List<OfficeDocumentAsset> Assets { get; } = new List<OfficeDocumentAsset>();
    }
}
