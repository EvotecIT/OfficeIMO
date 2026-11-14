namespace OfficeIMO.Email;

/// <summary>Extracts RFC 3156 PGP/MIME parts from the retained source bytes, and detects inline armored PGP text.</summary>
internal static class MimePgpExtractor {
    private const string CleartextMarker = "-----BEGIN PGP SIGNED MESSAGE-----";
    private const string MessageMarker = "-----BEGIN PGP MESSAGE-----";

    internal static EmailInlinePgpKind DetectInline(string? text) {
        if (string.IsNullOrEmpty(text)) return EmailInlinePgpKind.None;
        string trimmed = text!.TrimStart();
        if (trimmed.StartsWith(CleartextMarker, StringComparison.Ordinal)) return EmailInlinePgpKind.CleartextSigned;
        if (trimmed.StartsWith(MessageMarker, StringComparison.Ordinal)) return EmailInlinePgpKind.Message;
        return EmailInlinePgpKind.None;
    }

    /// <summary>Returns the exact bytes of the signed part and the decoded signature part of multipart/signed.</summary>
    internal static bool TryExtractSigned(EmailDocument document, long maximumBytes, List<EmailDiagnostic> diagnostics,
        out byte[]? signedEntity, out byte[]? signature, out string? micalg) {
        signedEntity = null;
        signature = null;
        micalg = null;
        if (!TrySplit(document, "multipart/signed", "pgp-signature", maximumBytes, diagnostics,
                out byte[]? source, out List<ArraySegment<byte>>? parts, out MimeValue? contentType)) return false;
        micalg = contentType!.GetParameter("micalg");
        ArraySegment<byte> signed = parts![0];
        if (signed.Count > maximumBytes) throw new InvalidDataException("The signed MIME part exceeds the configured input limit.");
        signedEntity = new byte[signed.Count];
        Buffer.BlockCopy(signed.Array!, signed.Offset, signedEntity, 0, signed.Count);
        signature = MimeSmimeExtractor.DecodePart(parts[1], source!, EmailReaderOptions.Default, diagnostics,
            "message/protection/signature", maximumBytes);
        return true;
    }

    /// <summary>Returns the decoded control part and encrypted payload of multipart/encrypted.</summary>
    internal static bool TryExtractEncrypted(EmailDocument document, long maximumBytes, List<EmailDiagnostic> diagnostics,
        out byte[]? control, out byte[]? payload) {
        control = null;
        payload = null;
        if (!TrySplit(document, "multipart/encrypted", "pgp-encrypted", maximumBytes, diagnostics,
                out byte[]? source, out List<ArraySegment<byte>>? parts, out _)) return false;
        control = MimeSmimeExtractor.DecodePart(parts![0], source!, EmailReaderOptions.Default, diagnostics,
            "message/protection/control", maximumBytes);
        payload = MimeSmimeExtractor.DecodePart(parts[1], source!, EmailReaderOptions.Default, diagnostics,
            "message/protection/payload", maximumBytes);
        return true;
    }

    /// <summary>RFC 3156 section 4: the control part must carry "Version: 1".</summary>
    internal static bool HasVersionOne(byte[] control) {
        string text = Encoding.ASCII.GetString(control);
        foreach (string line in text.Split('\n')) {
            string value = line.Trim();
            int colon = value.IndexOf(':');
            if (colon > 0 && value.Substring(0, colon).Trim().Equals("Version", StringComparison.OrdinalIgnoreCase) &&
                value.Substring(colon + 1).Trim() == "1") return true;
        }

        return false;
    }

    private static bool TrySplit(EmailDocument document, string expectedType, string expectedProtocol, long maximumBytes,
        List<EmailDiagnostic> diagnostics, out byte[]? source, out List<ArraySegment<byte>>? parts, out MimeValue? contentType) {
        source = document.Format == EmailFileFormat.Eml ? document.RawSource : null;
        parts = null;
        contentType = null;
        if (source == null) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_PGP_SOURCE_MISSING",
                "The PGP/MIME entity needs the original MIME bytes; only EML sources retain them.",
                EmailDiagnosticSeverity.Error, "message/protection"));
            return false;
        }

        if (source.LongLength > maximumBytes) throw new InvalidDataException("The message exceeds the configured input limit.");
        var headers = new List<EmailHeader>();
        int bodyOffset = MimeHeaderParser.Parse(source, 0, source.Length, EmailReaderOptions.Default, headers, diagnostics, "message");
        contentType = MimeValueParser.Parse(MimeHeaderParser.GetValue(headers, "Content-Type"), "text/plain", diagnostics, "message");
        string protocol = contentType.GetParameter("protocol") ?? string.Empty;
        string? boundary = contentType.GetParameter("boundary");
        if (!string.Equals(contentType.Value, expectedType, StringComparison.OrdinalIgnoreCase) ||
            protocol.IndexOf(expectedProtocol, StringComparison.OrdinalIgnoreCase) < 0 || boundary == null) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_PGP_ENTITY_INVALID",
                "The message is not a " + expectedType + " entity with an application/" + expectedProtocol + " protocol and boundary.",
                EmailDiagnosticSeverity.Error, "message/protection"));
            return false;
        }

        var state = new MimeParserState(EmailReaderOptions.Default, diagnostics, CancellationToken.None);
        parts = MimeParser.SplitMultipart(source, bodyOffset, source.Length - bodyOffset, boundary, state, "message");
        if (parts.Count != 2) {
            diagnostics.Add(new EmailDiagnostic("EMAIL_PGP_MULTIPART_INVALID",
                "An RFC 3156 " + expectedType + " entity must contain exactly two parts.",
                EmailDiagnosticSeverity.Error, "message/protection"));
            return false;
        }

        return true;
    }
}
