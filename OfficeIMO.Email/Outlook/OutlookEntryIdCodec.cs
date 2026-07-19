namespace OfficeIMO.Email;

/// <summary>Supported address-book EntryID family used by personal distribution-list members.</summary>
public enum OutlookEntryIdKind {
    /// <summary>Provider or shape is not recognized.</summary>
    Unknown = 0,
    /// <summary>MAPI One-Off EntryID.</summary>
    OneOff = 1,
    /// <summary>Wrapped contact Message EntryID.</summary>
    Contact = 2,
    /// <summary>Wrapped personal distribution-list Message EntryID.</summary>
    PersonalDistributionList = 3,
    /// <summary>Wrapped GAL mail-user EntryID.</summary>
    GalUser = 4,
    /// <summary>Wrapped GAL distribution-list EntryID.</summary>
    GalDistributionList = 5
}

/// <summary>Bounded codec for the One-Off and wrapped EntryIDs used by Outlook distribution lists.</summary>
public static class OutlookEntryIdCodec {
    private const int MaximumEntryIdBytes = 65_536;
    private static readonly byte[] OneOffProviderUid = {
        0x81, 0x2B, 0x1F, 0xA4, 0xBE, 0xA3, 0x10, 0x19,
        0x9D, 0x6E, 0x00, 0xDD, 0x01, 0x0F, 0x54, 0x02
    };
    private static readonly byte[] WrappedProviderUid = {
        0xC0, 0x91, 0xAD, 0xD3, 0x51, 0x9D, 0xCF, 0x11,
        0xA4, 0xA9, 0x00, 0xAA, 0x00, 0x47, 0xFA, 0xA4
    };

    static OutlookEntryIdCodec() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    /// <summary>Creates a Unicode MAPI One-Off EntryID.</summary>
    public static byte[] EncodeOneOff(EmailAddress address) {
        if (address == null) throw new ArgumentNullException(nameof(address));
        using (var stream = new MemoryStream()) {
            stream.Write(new byte[4], 0, 4);
            stream.Write(OneOffProviderUid, 0, OneOffProviderUid.Length);
            stream.WriteByte(0);
            stream.WriteByte(0);
            stream.WriteByte(0x09);
            stream.WriteByte(0xE8);
            WriteUnicodeString(stream, address.DisplayName ?? address.Address ?? string.Empty);
            WriteUnicodeString(stream, string.IsNullOrWhiteSpace(address.AddressType) ? "SMTP" : address.AddressType!);
            WriteUnicodeString(stream, address.Address ?? string.Empty);
            return stream.ToArray();
        }
    }

    /// <summary>Classifies a direct or wrapped distribution-list member EntryID.</summary>
    public static OutlookEntryIdKind Classify(byte[] entryId) {
        if (entryId == null) throw new ArgumentNullException(nameof(entryId));
        if (entryId.Length >= 24 && HasProvider(entryId, OneOffProviderUid)) return OutlookEntryIdKind.OneOff;
        if (entryId.Length < 22 || !HasProvider(entryId, WrappedProviderUid)) return OutlookEntryIdKind.Unknown;
        switch (entryId[20] & 0x0F) {
            case 0: return OutlookEntryIdKind.OneOff;
            case 3: return OutlookEntryIdKind.Contact;
            case 4: return OutlookEntryIdKind.PersonalDistributionList;
            case 5: return OutlookEntryIdKind.GalUser;
            case 6: return OutlookEntryIdKind.GalDistributionList;
            default: return OutlookEntryIdKind.Unknown;
        }
    }

    /// <summary>Attempts to decode a direct or type-zero wrapped One-Off EntryID.</summary>
    public static bool TryDecodeOneOff(byte[] entryId, out EmailAddress? address,
        out string? error, int ansiCodePage = 1252) {
        address = null;
        error = null;
        if (entryId == null) {
            error = "The EntryID is null.";
            return false;
        }
        if (entryId.Length > MaximumEntryIdBytes) {
            error = "The EntryID exceeds the supported bounded length.";
            return false;
        }
        int offset;
        if (entryId.Length >= 24 && HasProvider(entryId, OneOffProviderUid)) {
            offset = 0;
        } else if (entryId.Length >= 45 && HasProvider(entryId, WrappedProviderUid) &&
            (entryId[20] & 0x0F) == 0) {
            offset = 21;
            if (!HasProvider(entryId, OneOffProviderUid, offset + 4)) {
                error = "The wrapped EntryID does not contain a One-Off provider identifier.";
                return false;
            }
        } else {
            error = "The EntryID is not a supported direct or wrapped One-Off EntryID.";
            return false;
        }
        if (entryId.Length - offset < 24) {
            error = "The One-Off EntryID is truncated.";
            return false;
        }
        ushort version = ReadUInt16(entryId, offset + 20);
        ushort flags = ReadUInt16(entryId, offset + 22);
        if (version != 0) {
            error = "The One-Off EntryID version is unsupported.";
            return false;
        }
        bool unicode = (flags & 0x8000) != 0;
        Encoding encoding;
        try {
            encoding = unicode ? Encoding.Unicode : Encoding.GetEncoding(ansiCodePage,
                EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
        } catch (Exception exception) when (exception is ArgumentException || exception is NotSupportedException) {
            error = string.Concat("The ANSI One-Off code page is unavailable: ", exception.Message);
            return false;
        }
        int cursor = offset + 24;
        if (!TryReadTerminated(entryId, ref cursor, encoding, unicode ? 2 : 1, out string? displayName) ||
            !TryReadTerminated(entryId, ref cursor, encoding, unicode ? 2 : 1, out string? addressType) ||
            !TryReadTerminated(entryId, ref cursor, encoding, unicode ? 2 : 1, out string? emailAddress)) {
            error = "The One-Off EntryID contains a truncated or invalid string.";
            return false;
        }
        address = new EmailAddress(EmptyToNull(emailAddress), EmptyToNull(displayName)) {
            AddressType = EmptyToNull(addressType)
        };
        return true;
    }

    private static bool HasProvider(byte[] value, byte[] provider, int offset = 4) {
        if (offset < 0 || value.Length - offset < provider.Length) return false;
        for (int index = 0; index < provider.Length; index++) {
            if (value[offset + index] != provider[index]) return false;
        }
        return true;
    }

    private static bool TryReadTerminated(byte[] value, ref int cursor, Encoding encoding,
        int terminatorWidth, out string? text) {
        text = null;
        int start = cursor;
        while (cursor + terminatorWidth <= value.Length) {
            bool terminated = value[cursor] == 0 && (terminatorWidth == 1 || value[cursor + 1] == 0);
            if (terminated) {
                try {
                    text = encoding.GetString(value, start, cursor - start);
                } catch (DecoderFallbackException) {
                    return false;
                }
                cursor += terminatorWidth;
                return true;
            }
            cursor += terminatorWidth;
        }
        return false;
    }

    private static ushort ReadUInt16(byte[] value, int offset) =>
        (ushort)(value[offset] | (value[offset + 1] << 8));

    private static void WriteUnicodeString(Stream stream, string value) {
        byte[] bytes = Encoding.Unicode.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
        stream.WriteByte(0);
        stream.WriteByte(0);
    }

    private static string? EmptyToNull(string? value) => string.IsNullOrEmpty(value) ? null : value;
}
