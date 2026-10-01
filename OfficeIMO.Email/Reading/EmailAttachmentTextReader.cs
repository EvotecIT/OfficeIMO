namespace OfficeIMO.Email;

/// <summary>Decodes bounded text attachments using their MIME charset and any Unicode byte-order mark.</summary>
public static class EmailAttachmentTextReader {
    /// <summary>
    /// Reads decoded attachment bytes. A Unicode BOM takes precedence over the declared charset.
    /// Charsetless content uses the MIME reader's UTF-8 then Windows-1252 recovery policy.
    /// The attachment's reopenable source is disposed within the call.
    /// </summary>
    public static EmailAttachmentTextResult Read(EmailAttachment attachment,
        long maxBytes = 16L * 1024L * 1024L, CancellationToken cancellationToken = default) {
        if (attachment == null) throw new ArgumentNullException(nameof(attachment));
        if (maxBytes <= 0) throw new ArgumentOutOfRangeException(nameof(maxBytes));
        cancellationToken.ThrowIfCancellationRequested();
        using Stream source = attachment.OpenContentStream();
        using var output = new MemoryStream();
        byte[] buffer = new byte[8192];
        while (true) {
            cancellationToken.ThrowIfCancellationRequested();
            long remaining = maxBytes - output.Length;
            int count = source.Read(buffer, 0, remaining >= buffer.Length ? buffer.Length : (int)remaining + 1);
            if (count == 0) break;
            if (count > maxBytes - output.Length)
                throw new EmailLimitExceededException(nameof(maxBytes), output.Length + count, maxBytes);
            output.Write(buffer, 0, count);
        }
        cancellationToken.ThrowIfCancellationRequested();
        return Decode(output.ToArray(), attachment);
    }

    private static EmailAttachmentTextResult Decode(byte[] bytes, EmailAttachment attachment) {
        attachment.ContentTypeParameters.TryGetValue("charset", out string? charset);
        string location = "attachment/text";
        var diagnostics = new List<EmailDiagnostic>();
        int prefixLength = 0;
        string? bomCharset = null;
        if (bytes.Length >= 4 && bytes[0] == 0 && bytes[1] == 0 && bytes[2] == 0xfe && bytes[3] == 0xff) {
            bomCharset = "utf-32BE"; prefixLength = 4;
        } else if (bytes.Length >= 4 && bytes[0] == 0xff && bytes[1] == 0xfe && bytes[2] == 0 && bytes[3] == 0) {
            bomCharset = "utf-32"; prefixLength = 4;
        } else if (bytes.Length >= 3 && bytes[0] == 0xef && bytes[1] == 0xbb && bytes[2] == 0xbf) {
            bomCharset = "utf-8"; prefixLength = 3;
        } else if (bytes.Length >= 2 && bytes[0] == 0xff && bytes[1] == 0xfe) {
            bomCharset = "utf-16"; prefixLength = 2;
        } else if (bytes.Length >= 2 && bytes[0] == 0xfe && bytes[1] == 0xff) {
            bomCharset = "utf-16BE"; prefixLength = 2;
        }
        if (bomCharset != null) {
            if (!string.IsNullOrWhiteSpace(charset)) {
                bool matches;
                try { matches = MimeTextCodec.ResolveStrictEncoding(charset).CodePage == MimeTextCodec.ResolveStrictEncoding(bomCharset).CodePage; }
                catch (Exception exception) when (exception is ArgumentException || exception is NotSupportedException) { matches = false; }
                if (!matches) diagnostics.Add(new EmailDiagnostic("EMAIL_ATTACHMENT_BOM_OVERRIDES_CHARSET",
                    "The Unicode byte-order mark takes precedence over the declared MIME charset.", EmailDiagnosticSeverity.Warning, location));
            }
            charset = bomCharset;
        }
        byte[] payload = bytes;
        if (prefixLength != 0) {
            payload = new byte[bytes.Length - prefixLength];
            Buffer.BlockCopy(bytes, prefixLength, payload, 0, payload.Length);
        }
        if (!string.IsNullOrWhiteSpace(charset)) {
            try {
                string strictText = MimeTextCodec.ResolveStrictEncoding(charset).GetString(payload);
                return new EmailAttachmentTextResult(strictText, charset!, bytes.LongLength, diagnostics.AsReadOnly());
            } catch (DecoderFallbackException) {
                diagnostics.Add(new EmailDiagnostic("EMAIL_ATTACHMENT_TEXT_INVALID_ENCODING",
                    "The attachment contains bytes invalid for its charset; replacement decoding was used.", EmailDiagnosticSeverity.Warning, location));
            } catch (Exception exception) when (exception is ArgumentException || exception is NotSupportedException) {
                // The shared MIME decoder reports the unavailable charset and its fallback below.
            }
        }
        string text = MimeTextCodec.DecodeText(payload, charset, diagnostics, location);
        string effectiveCharset = diagnostics.Any(item => item.Code == "EMAIL_MIME_CHARSET_UNSUPPORTED") ? "utf-8" :
            diagnostics.Any(item => item.Code == "EMAIL_MIME_CHARSET_GUESSED") ? "windows-1252" : charset ?? "utf-8";
        return new EmailAttachmentTextResult(text, effectiveCharset, bytes.LongLength, diagnostics.AsReadOnly());
    }
}

/// <summary>Text extraction and decoding evidence for one attachment.</summary>
public sealed class EmailAttachmentTextResult {
    internal EmailAttachmentTextResult(string text, string charset, long bytesRead, IReadOnlyList<EmailDiagnostic> diagnostics) {
        Text = text; Charset = charset; BytesRead = bytesRead; Diagnostics = diagnostics;
    }
    /// <summary>Decoded text without a Unicode byte-order mark.</summary>
    public string Text { get; }
    /// <summary>Effective charset used to decode the content.</summary>
    public string Charset { get; }
    /// <summary>Number of decoded attachment bytes consumed before text decoding.</summary>
    public long BytesRead { get; }
    /// <summary>Warnings about recovery or conflicting encoding declarations.</summary>
    public IReadOnlyList<EmailDiagnostic> Diagnostics { get; }
}
