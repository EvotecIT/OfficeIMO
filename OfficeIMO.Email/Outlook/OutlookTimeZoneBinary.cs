namespace OfficeIMO.Email;

/// <summary>MS-OXOCAL time-zone definition and legacy structure reader/writer.</summary>
public static class OutlookTimeZoneBinary {
    private const int TimeZoneRuleSize = 66;
    private const int MaximumRules = 512;

    /// <summary>Decodes a TZDEFINITION while retaining unsupported or malformed input losslessly.</summary>
    public static OutlookTimeZoneDefinition DecodeDefinition(byte[] bytes) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        var definition = new OutlookTimeZoneDefinition();
        try {
            var cursor = new Cursor(bytes);
            byte major = cursor.ReadByte();
            byte minor = cursor.ReadByte();
            if (major != 0x02 || minor != 0x01)
                throw new NotSupportedException("Unsupported Outlook time-zone definition version.");
            ushort headerSize = cursor.ReadUInt16();
            int headerEnd = checked(4 + headerSize);
            if (headerSize < 6 || headerEnd > bytes.Length)
                throw new InvalidDataException("Invalid Outlook time-zone definition header size.");
            ushort reserved = cursor.ReadUInt16();
            if (reserved != 0x0002)
                throw new InvalidDataException("Invalid Outlook time-zone definition header marker.");
            ushort keyLength = cursor.ReadUInt16();
            int keyBytes = checked(keyLength * 2);
            if (headerSize != checked(6 + keyBytes))
                throw new InvalidDataException("Outlook time-zone key length does not match the header size.");
            definition.KeyName = Encoding.Unicode.GetString(cursor.ReadBytes(keyBytes));
            ushort ruleCount = cursor.ReadUInt16();
            if (ruleCount == 0 || ruleCount > MaximumRules)
                throw new InvalidDataException("Invalid Outlook time-zone rule count.");
            if (cursor.Position != headerEnd)
                throw new InvalidDataException("Outlook time-zone header fields do not match cbHeader.");
            if (cursor.Remaining != checked(ruleCount * TimeZoneRuleSize))
                throw new InvalidDataException("Outlook time-zone rule data has an invalid size.");
            for (int index = 0; index < ruleCount; index++) definition.Rules.Add(ReadRule(cursor));
            if (!cursor.AtEnd) throw new InvalidDataException("The time-zone definition contains trailing data.");
            definition.AcceptDecodedState(bytes);
        } catch (Exception ex) when (IsDecodeException(ex)) {
            definition.SetDecodeFailure(bytes, ex.Message);
        }
        return definition;
    }

    /// <summary>Encodes a TZDEFINITION, preserving untouched decoded source bytes exactly.</summary>
    public static byte[] EncodeDefinition(OutlookTimeZoneDefinition definition) {
        if (definition == null) throw new ArgumentNullException(nameof(definition));
        if (definition.CanPreserveRawState) return (byte[])definition.RawState!.Clone();
        ValidateDefinition(definition);
        using (var stream = new MemoryStream())
        using (var writer = new BinaryWriter(stream, Encoding.Unicode, true)) {
            writer.Write((byte)0x02);
            writer.Write((byte)0x01);
            int keyLength = definition.KeyName!.Length;
            writer.Write(checked((ushort)(6 + keyLength * 2)));
            writer.Write((ushort)0x0002);
            writer.Write(checked((ushort)keyLength));
            writer.Write(Encoding.Unicode.GetBytes(definition.KeyName));
            writer.Write(checked((ushort)definition.Rules.Count));
            foreach (OutlookTimeZoneRule rule in definition.Rules) WriteRule(writer, rule);
            writer.Flush();
            return stream.ToArray();
        }
    }

    /// <summary>Decodes PidLidTimeZoneStruct while retaining malformed input losslessly.</summary>
    public static OutlookTimeZoneStructure DecodeStructure(byte[] bytes) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        var structure = new OutlookTimeZoneStructure();
        try {
            if (bytes.Length != 48)
                throw new InvalidDataException("PidLidTimeZoneStruct must contain exactly 48 bytes.");
            var cursor = new Cursor(bytes);
            structure.Rule.BiasMinutes = cursor.ReadInt32();
            structure.Rule.StandardBiasMinutes = cursor.ReadInt32();
            structure.Rule.DaylightBiasMinutes = cursor.ReadInt32();
            structure.StandardYear = cursor.ReadUInt16();
            structure.Rule.StandardTransition = ReadTransition(cursor);
            structure.DaylightYear = cursor.ReadUInt16();
            structure.Rule.DaylightTransition = ReadTransition(cursor);
            if (!cursor.AtEnd) throw new InvalidDataException("PidLidTimeZoneStruct contains trailing data.");
            structure.AcceptDecodedState(bytes);
        } catch (Exception ex) when (IsDecodeException(ex)) {
            structure.SetDecodeFailure(bytes, ex.Message);
        }
        return structure;
    }

    /// <summary>Encodes PidLidTimeZoneStruct, preserving untouched decoded source bytes exactly.</summary>
    public static byte[] EncodeStructure(OutlookTimeZoneStructure structure) {
        if (structure == null) throw new ArgumentNullException(nameof(structure));
        if (structure.CanPreserveRawState) return (byte[])structure.RawState!.Clone();
        ValidateRule(structure.Rule);
        using (var stream = new MemoryStream(48))
        using (var writer = new BinaryWriter(stream, Encoding.Unicode, true)) {
            writer.Write(structure.Rule.BiasMinutes);
            writer.Write(structure.Rule.StandardBiasMinutes);
            writer.Write(structure.Rule.DaylightBiasMinutes);
            writer.Write(structure.StandardYear);
            WriteTransition(writer, structure.Rule.StandardTransition);
            writer.Write(structure.DaylightYear);
            WriteTransition(writer, structure.Rule.DaylightTransition);
            writer.Flush();
            return stream.ToArray();
        }
    }

    private static OutlookTimeZoneRule ReadRule(Cursor cursor) {
        byte major = cursor.ReadByte();
        byte minor = cursor.ReadByte();
        if (major != 0x02 || minor != 0x01)
            throw new NotSupportedException("Unsupported Outlook TZRule version.");
        if (cursor.ReadUInt16() != 0x003E)
            throw new InvalidDataException("Invalid Outlook TZRule marker.");
        var rule = new OutlookTimeZoneRule {
            Flags = cursor.ReadUInt16(),
            EffectiveYear = cursor.ReadUInt16()
        };
        byte[] unused = cursor.ReadBytes(14);
        if (unused.Any(value => value != 0))
            throw new InvalidDataException("Outlook TZRule reserved bytes are not zero.");
        rule.BiasMinutes = cursor.ReadInt32();
        rule.StandardBiasMinutes = cursor.ReadInt32();
        rule.DaylightBiasMinutes = cursor.ReadInt32();
        rule.StandardTransition = ReadTransition(cursor);
        rule.DaylightTransition = ReadTransition(cursor);
        return rule;
    }

    private static void WriteRule(BinaryWriter writer, OutlookTimeZoneRule rule) {
        ValidateRule(rule);
        writer.Write((byte)0x02);
        writer.Write((byte)0x01);
        writer.Write((ushort)0x003E);
        writer.Write(rule.Flags);
        writer.Write(rule.EffectiveYear);
        writer.Write(new byte[14]);
        writer.Write(rule.BiasMinutes);
        writer.Write(rule.StandardBiasMinutes);
        writer.Write(rule.DaylightBiasMinutes);
        WriteTransition(writer, rule.StandardTransition);
        WriteTransition(writer, rule.DaylightTransition);
    }

    private static OutlookTimeZoneTransition ReadTransition(Cursor cursor) {
        var transition = new OutlookTimeZoneTransition(cursor.ReadUInt16(), cursor.ReadUInt16(),
            cursor.ReadUInt16(), cursor.ReadUInt16(), cursor.ReadUInt16(), cursor.ReadUInt16(),
            cursor.ReadUInt16(), cursor.ReadUInt16());
        ValidateTransition(transition);
        return transition;
    }

    private static void WriteTransition(BinaryWriter writer, OutlookTimeZoneTransition transition) {
        ValidateTransition(transition);
        writer.Write(transition.Year);
        writer.Write(transition.Month);
        writer.Write(transition.DayOfWeek);
        writer.Write(transition.Day);
        writer.Write(transition.Hour);
        writer.Write(transition.Minute);
        writer.Write(transition.Second);
        writer.Write(transition.Milliseconds);
    }

    private static void ValidateDefinition(OutlookTimeZoneDefinition definition) {
        if (string.IsNullOrEmpty(definition.KeyName))
            throw new InvalidOperationException("An Outlook time-zone definition requires KeyName.");
        if (definition.KeyName!.Length > (ushort.MaxValue - 6) / 2)
            throw new ArgumentOutOfRangeException(nameof(definition), "The Outlook time-zone key is too long.");
        if (definition.Rules.Count == 0 || definition.Rules.Count > MaximumRules)
            throw new InvalidOperationException("An Outlook time-zone definition requires a bounded rule set.");
        int activeRules = definition.Rules.Count(rule => (rule.Flags & 0x0002) != 0);
        if (activeRules != 1)
            throw new InvalidOperationException("Exactly one Outlook time-zone rule must be effective.");
        foreach (OutlookTimeZoneRule rule in definition.Rules) ValidateRule(rule);
    }

    private static void ValidateRule(OutlookTimeZoneRule rule) {
        if (rule == null) throw new ArgumentNullException(nameof(rule));
        ValidateTransition(rule.StandardTransition);
        ValidateTransition(rule.DaylightTransition);
        if (rule.StandardTransition.IsDisabled != rule.DaylightTransition.IsDisabled)
            throw new InvalidOperationException("Standard and daylight transitions must both be present or absent.");
    }

    private static void ValidateTransition(OutlookTimeZoneTransition transition) {
        if (transition == null) throw new ArgumentNullException(nameof(transition));
        if (transition.IsDisabled) {
            if (transition.Year != 0 || transition.DayOfWeek != 0 || transition.Day != 0 || transition.Hour != 0 ||
                transition.Minute != 0 || transition.Second != 0 || transition.Milliseconds != 0)
                throw new InvalidDataException("A disabled Outlook time-zone transition must be zero-filled.");
            return;
        }
        if (transition.Day == 0)
            throw new InvalidDataException("An Outlook time-zone transition requires a day.");
        if (transition.Year == 0 && transition.Day > 5)
            throw new InvalidDataException("A relative Outlook time-zone transition occurrence must be 1 through 5.");
        if (transition.Year != 0 && transition.Day > DateTime.DaysInMonth(transition.Year, transition.Month))
            throw new InvalidDataException("An absolute Outlook time-zone transition has an invalid date.");
    }

    private static bool IsDecodeException(Exception ex) => ex is InvalidDataException ||
        ex is NotSupportedException || ex is ArgumentException || ex is OverflowException;

    private sealed class Cursor {
        private readonly byte[] _bytes;
        internal Cursor(byte[] bytes) => _bytes = bytes;
        internal int Position { get; private set; }
        internal int Remaining => _bytes.Length - Position;
        internal bool AtEnd => Position == _bytes.Length;
        internal byte ReadByte() { Ensure(1); return _bytes[Position++]; }
        internal ushort ReadUInt16() { Ensure(2); ushort value = BitConverter.ToUInt16(_bytes, Position); Position += 2; return value; }
        internal int ReadInt32() { Ensure(4); int value = BitConverter.ToInt32(_bytes, Position); Position += 4; return value; }
        internal byte[] ReadBytes(int count) { Ensure(count); var result = new byte[count]; Buffer.BlockCopy(_bytes, Position, result, 0, count); Position += count; return result; }
        private void Ensure(int count) { if (count < 0 || count > Remaining) throw new InvalidDataException("The Outlook time-zone BLOB is truncated."); }
    }
}
