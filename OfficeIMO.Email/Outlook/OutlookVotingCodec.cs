namespace OfficeIMO.Email;

/// <summary>MS-OXOMSG PidLidVerbStream codec.</summary>
internal static class OutlookVotingCodec {
    private const ushort Version = 0x0102;
    private const ushort ExtrasVersion = 0x0104;
    private const uint VotingVerbType = 4;
    private const int MaximumOptions = 255;

    static OutlookVotingCodec() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    internal static bool TryDecode(byte[] value, int codePage, IList<OutlookVoteOption> options,
        out string? error) {
        if (value == null) throw new ArgumentNullException(nameof(value));
        if (options == null) throw new ArgumentNullException(nameof(options));
        options.Clear();
        try {
            using (var stream = new MemoryStream(value, writable: false))
            using (var reader = new BinaryReader(stream, Encoding.UTF8, leaveOpen: true)) {
                if (reader.ReadUInt16() != Version) throw new InvalidDataException("Unsupported verb-stream version.");
                uint count = reader.ReadUInt32();
                if (count > MaximumOptions) throw new InvalidDataException("Voting option count exceeds 255.");

                Encoding ansi = StrictEncoding(codePage);
                for (int index = 0; index < (int)count; index++) options.Add(ReadOption(reader, ansi));
                if (reader.ReadUInt16() != ExtrasVersion) {
                    throw new InvalidDataException("Voting option extras marker is missing.");
                }
                for (int index = 0; index < options.Count; index++) {
                    string displayName = ReadUnicodeByteCounted(reader);
                    string repeated = ReadUnicodeByteCounted(reader);
                    if (!string.Equals(displayName, repeated, StringComparison.Ordinal)) {
                        throw new InvalidDataException("Voting option Unicode display-name copies differ.");
                    }
                    if (!string.IsNullOrEmpty(displayName)) options[index].DisplayName = displayName;
                }
                if (stream.Position != stream.Length) {
                    throw new InvalidDataException("Voting verb stream contains trailing bytes.");
                }
            }
            error = null;
            return true;
        } catch (Exception exception) when (exception is EndOfStreamException ||
            exception is InvalidDataException || exception is DecoderFallbackException ||
            exception is ArgumentException || exception is NotSupportedException) {
            options.Clear();
            error = exception.Message;
            return false;
        }
    }

    internal static bool TryEncode(IReadOnlyList<OutlookVoteOption> options, int codePage,
        out byte[]? value, out string? error) {
        if (options == null) throw new ArgumentNullException(nameof(options));
        try {
            if (options.Count > MaximumOptions) throw new InvalidDataException("Voting option count exceeds 255.");
            Encoding ansi = StrictEncoding(codePage);
            using (var stream = new MemoryStream())
            using (var writer = new BinaryWriter(stream, Encoding.UTF8, leaveOpen: true)) {
                writer.Write(Version);
                writer.Write((uint)options.Count);
                for (int index = 0; index < options.Count; index++) WriteOption(writer, ansi, options[index], index);
                writer.Write(ExtrasVersion);
                for (int index = 0; index < options.Count; index++) {
                    WriteUnicodeByteCounted(writer, options[index].DisplayName);
                    WriteUnicodeByteCounted(writer, options[index].DisplayName);
                }
                writer.Flush();
                value = stream.ToArray();
            }
            error = null;
            return true;
        } catch (Exception exception) when (exception is InvalidDataException ||
            exception is EncoderFallbackException || exception is ArgumentException ||
            exception is NotSupportedException) {
            value = null;
            error = exception.Message;
            return false;
        }
    }

    private static OutlookVoteOption ReadOption(BinaryReader reader, Encoding ansi) {
        if (reader.ReadUInt32() != VotingVerbType) throw new InvalidDataException("Unsupported voting verb type.");
        string displayName = ReadAnsiByteCounted(reader, ansi);
        string messageClass = ReadAnsiByteCounted(reader, ansi);
        if (!string.Equals(messageClass, "IPM.Note", StringComparison.Ordinal)) {
            throw new InvalidDataException("Voting option message class is not IPM.Note.");
        }
        if (ReadAnsiByteCounted(reader, ansi).Length != 0) {
            throw new InvalidDataException("Voting option internal string is not empty.");
        }
        string repeated = ReadAnsiByteCounted(reader, ansi);
        if (!string.Equals(displayName, repeated, StringComparison.Ordinal)) {
            throw new InvalidDataException("Voting option ANSI display-name copies differ.");
        }
        if (reader.ReadUInt32() != 0 || reader.ReadByte() != 0) {
            throw new InvalidDataException("Voting option reserved fields are invalid.");
        }
        uint useUsHeaders = reader.ReadUInt32();
        if (useUsHeaders > 1) throw new InvalidDataException("Voting option reply-header flag is invalid.");
        if (reader.ReadUInt32() != 1) throw new InvalidDataException("Voting option reserved field is invalid.");
        uint sendBehavior = reader.ReadUInt32();
        if (sendBehavior != (uint)OutlookVoteSendBehavior.Automatic &&
            sendBehavior != (uint)OutlookVoteSendBehavior.Prompt) {
            throw new InvalidDataException("Voting option send behavior is invalid.");
        }
        if (reader.ReadUInt32() != 2) throw new InvalidDataException("Voting option reserved field is invalid.");
        uint id = reader.ReadUInt32();
        if (id > int.MaxValue) throw new InvalidDataException("Voting option ID is out of range.");
        if (reader.ReadUInt32() != uint.MaxValue) throw new InvalidDataException("Voting option terminator is invalid.");
        return new OutlookVoteOption(displayName) {
            Id = (int)id,
            SendBehavior = (OutlookVoteSendBehavior)sendBehavior,
            UseUsReplyHeaders = useUsHeaders == 1
        };
    }

    private static void WriteOption(BinaryWriter writer, Encoding ansi, OutlookVoteOption option, int index) {
        if (option == null) throw new InvalidDataException("A voting option cannot be null.");
        writer.Write(VotingVerbType);
        WriteAnsiByteCounted(writer, ansi, option.DisplayName);
        WriteAnsiByteCounted(writer, ansi, "IPM.Note");
        writer.Write((byte)0);
        WriteAnsiByteCounted(writer, ansi, option.DisplayName);
        writer.Write(0U);
        writer.Write((byte)0);
        writer.Write(option.UseUsReplyHeaders ? 1U : 0U);
        writer.Write(1U);
        uint sendBehavior = (uint)option.SendBehavior;
        if (sendBehavior != 1 && sendBehavior != 2) throw new InvalidDataException("Voting option send behavior is invalid.");
        writer.Write(sendBehavior);
        writer.Write(2U);
        int id = option.Id > 0 ? option.Id : index + 1;
        writer.Write((uint)id);
        writer.Write(uint.MaxValue);
    }

    private static string ReadAnsiByteCounted(BinaryReader reader, Encoding encoding) {
        int count = reader.ReadByte();
        byte[] bytes = reader.ReadBytes(count);
        if (bytes.Length != count) throw new EndOfStreamException();
        return encoding.GetString(bytes);
    }

    private static void WriteAnsiByteCounted(BinaryWriter writer, Encoding encoding, string value) {
        if (value == null) throw new InvalidDataException("A voting option string cannot be null.");
        byte[] bytes = encoding.GetBytes(value);
        if (bytes.Length > byte.MaxValue) throw new InvalidDataException("A voting option string exceeds 255 ANSI bytes.");
        writer.Write((byte)bytes.Length);
        writer.Write(bytes);
    }

    private static string ReadUnicodeByteCounted(BinaryReader reader) {
        int characterCount = reader.ReadByte();
        byte[] bytes = reader.ReadBytes(checked(characterCount * 2));
        if (bytes.Length != characterCount * 2) throw new EndOfStreamException();
        return new UnicodeEncoding(false, false, true).GetString(bytes);
    }

    private static void WriteUnicodeByteCounted(BinaryWriter writer, string value) {
        if (value == null) throw new InvalidDataException("A voting option string cannot be null.");
        if (value.Length > byte.MaxValue) throw new InvalidDataException("A voting option string exceeds 255 Unicode characters.");
        writer.Write((byte)value.Length);
        writer.Write(new UnicodeEncoding(false, false, true).GetBytes(value));
    }

    private static Encoding StrictEncoding(int codePage) => Encoding.GetEncoding(
        codePage > 0 ? codePage : 1252,
        EncoderFallback.ExceptionFallback,
        DecoderFallback.ExceptionFallback);
}
