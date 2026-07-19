namespace OfficeIMO.Email;

/// <summary>MS-OXOCAL recurrence BLOB reader and writer.</summary>
public static class OutlookRecurrenceBinary {
    private static readonly DateTime ReferenceDate = new DateTime(1601, 1, 1, 0, 0, 0, DateTimeKind.Unspecified);
    private const uint NoEndDate = 0x5AE980DF;
    private const int MaximumInstances = 100000;

    static OutlookRecurrenceBinary() {
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
    }

    /// <summary>Decodes an AppointmentRecurrencePattern while retaining invalid input for lossless round trips.</summary>
    public static OutlookRecurrence DecodeAppointment(byte[] bytes, int string8CodePage = 1252) =>
        Decode(bytes, appointment: true, string8CodePage);

    /// <summary>Decodes a task RecurrencePattern while retaining invalid input for lossless round trips.</summary>
    public static OutlookRecurrence DecodeTask(byte[] bytes) => Decode(bytes, appointment: false, 1252);

    /// <summary>Encodes an AppointmentRecurrencePattern, preserving untouched decoded source bytes exactly.</summary>
    public static byte[] EncodeAppointment(OutlookRecurrence recurrence, int string8CodePage = 1252) =>
        Encode(recurrence, appointment: true, string8CodePage);

    /// <summary>Encodes a task RecurrencePattern, preserving untouched decoded source bytes exactly.</summary>
    public static byte[] EncodeTask(OutlookRecurrence recurrence) => Encode(recurrence, appointment: false, 1252);

    private static OutlookRecurrence Decode(byte[] bytes, bool appointment, int codePage) {
        if (bytes == null) throw new ArgumentNullException(nameof(bytes));
        var recurrence = new OutlookRecurrence();
        try {
            var cursor = new Cursor(bytes);
            ParsedPattern parsed = ReadPattern(cursor, recurrence);
            if (appointment) ReadAppointmentTail(cursor, recurrence, parsed.ModifiedCount, codePage);
            if (!cursor.AtEnd) throw new InvalidDataException("The recurrence BLOB contains trailing data.");
            recurrence.AcceptDecodedState(bytes);
        } catch (Exception ex) when (ex is InvalidDataException || ex is NotSupportedException ||
                                      ex is ArgumentOutOfRangeException || ex is OverflowException ||
                                      ex is ArgumentException) {
            recurrence.SetDecodeFailure(bytes, ex.Message);
        }
        return recurrence;
    }

    private static ParsedPattern ReadPattern(Cursor cursor, OutlookRecurrence recurrence) {
        if (cursor.ReadUInt16() != 0x3004 || cursor.ReadUInt16() != 0x3004)
            throw new InvalidDataException("Unsupported RecurrencePattern reader or writer version.");
        ushort frequency = cursor.ReadUInt16();
        ushort patternType = cursor.ReadUInt16();
        ushort calendarType = cursor.ReadUInt16();
        cursor.ReadUInt32(); // FirstDateTime is derivable from the pattern and retained in RawState.
        uint period = cursor.ReadUInt32();
        uint sliding = cursor.ReadUInt32();

        recurrence.Frequency = DecodeFrequency(frequency);
        recurrence.PatternKind = DecodePatternKind(patternType);
        if (!IsValidFrequencyPattern(recurrence.Frequency, recurrence.PatternKind))
            throw new InvalidDataException("The recurrence frequency and pattern type are incompatible.");
        recurrence.CalendarType = calendarType;
        recurrence.Sliding = sliding != 0;
        recurrence.Interval = DecodeInterval(recurrence.Frequency, period);
        switch (patternType) {
            case 0x0000: break;
            case 0x0001:
                recurrence.DaysOfWeek = DecodeDays(cursor.ReadUInt32());
                break;
            case 0x0002:
            case 0x0004:
                recurrence.DayOfMonth = checked((int)cursor.ReadUInt32());
                break;
            case 0x0003:
                recurrence.DaysOfWeek = DecodeDays(cursor.ReadUInt32());
                uint ordinal = cursor.ReadUInt32();
                if (ordinal < 1 || ordinal > 5) throw new InvalidDataException("Invalid MonthNth ordinal.");
                recurrence.WeekOrdinal = (OutlookRecurrenceWeekOrdinal)ordinal;
                break;
            default:
                throw new NotSupportedException("The recurrence uses a non-Gregorian or unsupported pattern type.");
        }

        uint endType = cursor.ReadUInt32();
        uint occurrenceCount = cursor.ReadUInt32();
        uint firstDay = cursor.ReadUInt32();
        if (firstDay > 6) throw new InvalidDataException("Invalid first day of week.");
        recurrence.FirstDayOfWeek = (DayOfWeek)firstDay;
        recurrence.RangeKind = DecodeRange(endType);
        if (recurrence.RangeKind == OutlookRecurrenceRangeKind.OccurrenceCount) {
            if (occurrenceCount == 0 || occurrenceCount > int.MaxValue)
                throw new InvalidDataException("Invalid recurrence occurrence count.");
            recurrence.OccurrenceCount = (int)occurrenceCount;
        }

        uint deletedCount = cursor.ReadCount(MaximumInstances, "deleted recurrence instances");
        for (uint index = 0; index < deletedCount; index++)
            recurrence.DeletedOccurrenceDates.Add(FromMinutes(cursor.ReadUInt32()).Date);
        uint modifiedCount = cursor.ReadCount(MaximumInstances, "modified recurrence instances");
        if (modifiedCount > deletedCount) throw new InvalidDataException("Modified instance count exceeds deleted count.");
        var modifiedDates = new List<DateTime>((int)modifiedCount);
        for (uint index = 0; index < modifiedCount; index++)
            modifiedDates.Add(FromMinutes(cursor.ReadUInt32()).Date);
        DateTime startDate = FromMinutes(cursor.ReadUInt32()).Date;
        uint rawEndDate = cursor.ReadUInt32();
        recurrence.Start = startDate;
        if (recurrence.RangeKind == OutlookRecurrenceRangeKind.EndDate) recurrence.EndDate = FromMinutes(rawEndDate).Date;
        else if (recurrence.RangeKind != OutlookRecurrenceRangeKind.NoEnd && rawEndDate != NoEndDate)
            recurrence.EndDate = FromMinutes(rawEndDate).Date;
        return new ParsedPattern(modifiedCount, modifiedDates);
    }

    private static void ReadAppointmentTail(Cursor cursor, OutlookRecurrence recurrence, uint modifiedCount,
        int codePage) {
        uint readerVersion = cursor.ReadUInt32();
        uint writerVersion = cursor.ReadUInt32();
        if (readerVersion != 0x00003006 || (writerVersion != 0x00003008 && writerVersion != 0x00003009))
            throw new InvalidDataException("Unsupported AppointmentRecurrencePattern extension version.");
        uint startOffset = cursor.ReadUInt32();
        uint endOffset = cursor.ReadUInt32();
        if (endOffset < startOffset) throw new InvalidDataException("Appointment recurrence end offset precedes its start offset.");
        recurrence.Start = recurrence.Start.Date.AddMinutes(startOffset);
        recurrence.Duration = TimeSpan.FromMinutes(endOffset - startOffset);
        ushort exceptionCount = cursor.ReadUInt16();
        if (exceptionCount != modifiedCount)
            throw new InvalidDataException("Appointment exception count does not match modified instances.");

        Encoding ansi = GetEncoding(codePage);
        var parsedExceptions = new List<ParsedException>(exceptionCount);
        for (int index = 0; index < exceptionCount; index++) parsedExceptions.Add(ReadException(cursor, ansi));
        cursor.SkipCountedBlock("ReservedBlock1");
        for (int index = 0; index < parsedExceptions.Count; index++)
            ReadExtendedException(cursor, parsedExceptions[index], writerVersion);
        cursor.SkipCountedBlock("ReservedBlock2");
        foreach (ParsedException parsed in parsedExceptions) recurrence.Exceptions.Add(parsed.Exception);
    }

    private static ParsedException ReadException(Cursor cursor, Encoding ansi) {
        var exception = new OutlookRecurrenceException {
            Start = FromMinutes(cursor.ReadUInt32()),
            End = FromMinutes(cursor.ReadUInt32()),
            OriginalStart = FromMinutes(cursor.ReadUInt32())
        };
        ushort flags = cursor.ReadUInt16();
        if ((flags & 0x0001) != 0) exception.Subject = cursor.ReadAnsiWithTwoLengths(ansi, "subject");
        if ((flags & 0x0002) != 0) exception.MeetingType = ToInt32(cursor.ReadUInt32());
        if ((flags & 0x0004) != 0) exception.ReminderDeltaMinutes = ToInt32(cursor.ReadUInt32());
        if ((flags & 0x0008) != 0) exception.ReminderIsSet = cursor.ReadUInt32() != 0;
        if ((flags & 0x0010) != 0) exception.Location = cursor.ReadAnsiWithTwoLengths(ansi, "location");
        if ((flags & 0x0020) != 0) exception.BusyStatus = ToInt32(cursor.ReadUInt32());
        if ((flags & 0x0040) != 0) exception.HasAttachments = cursor.ReadUInt32() != 0;
        if ((flags & 0x0080) != 0) exception.IsAllDay = cursor.ReadUInt32() != 0;
        if ((flags & 0x0100) != 0) exception.AppointmentColor = ToInt32(cursor.ReadUInt32());
        exception.HasExceptionalBody = (flags & 0x0200) != 0;
        if ((flags & 0xFC00) != 0) throw new InvalidDataException("Appointment exception contains unknown override flags.");
        return new ParsedException(exception, flags);
    }

    private static void ReadExtendedException(Cursor cursor, ParsedException parsed, uint writerVersion) {
        if (writerVersion >= 0x00003009) cursor.SkipCountedBlock("ChangeHighlight");
        cursor.SkipCountedBlock("ReservedBlockEE1");
        if ((parsed.Flags & 0x0011) != 0) {
            parsed.Exception.Start = FromMinutes(cursor.ReadUInt32());
            parsed.Exception.End = FromMinutes(cursor.ReadUInt32());
            parsed.Exception.OriginalStart = FromMinutes(cursor.ReadUInt32());
            if ((parsed.Flags & 0x0001) != 0) parsed.Exception.Subject = cursor.ReadUnicode16("subject");
            if ((parsed.Flags & 0x0010) != 0) parsed.Exception.Location = cursor.ReadUnicode16("location");
            cursor.SkipCountedBlock("ReservedBlockEE2");
        }
    }

    private static byte[] Encode(OutlookRecurrence recurrence, bool appointment, int codePage) {
        if (recurrence == null) throw new ArgumentNullException(nameof(recurrence));
        if (recurrence.CanPreserveRawState) return (byte[])recurrence.RawState!.Clone();
        ValidateForWrite(recurrence, appointment);
        Encoding ansi = GetEncoding(codePage);
        using (var stream = new MemoryStream())
        using (var writer = new BinaryWriter(stream, Encoding.UTF8, true)) {
            WritePattern(writer, recurrence);
            if (appointment) WriteAppointmentTail(writer, recurrence, ansi);
            writer.Flush();
            return stream.ToArray();
        }
    }

    private static void WritePattern(BinaryWriter writer, OutlookRecurrence recurrence) {
        writer.Write((ushort)0x3004);
        writer.Write((ushort)0x3004);
        writer.Write(EncodeFrequency(recurrence.Frequency));
        writer.Write(EncodePatternKind(recurrence.PatternKind));
        writer.Write(recurrence.CalendarType);
        uint period = EncodePeriod(recurrence);
        writer.Write(ComputeFirstDateTime(recurrence, period));
        writer.Write(period);
        writer.Write(recurrence.Sliding ? 1U : 0U);
        switch (recurrence.PatternKind) {
            case OutlookRecurrencePatternKind.Day: break;
            case OutlookRecurrencePatternKind.Week:
                writer.Write((uint)recurrence.DaysOfWeek);
                break;
            case OutlookRecurrencePatternKind.MonthDay:
                writer.Write((uint)recurrence.DayOfMonth!.Value);
                break;
            case OutlookRecurrencePatternKind.MonthEnd:
                writer.Write(31U);
                break;
            case OutlookRecurrencePatternKind.MonthNth:
                writer.Write((uint)recurrence.DaysOfWeek);
                writer.Write((uint)recurrence.WeekOrdinal!.Value);
                break;
        }
        writer.Write(EncodeRange(recurrence.RangeKind));
        writer.Write((uint)(recurrence.OccurrenceCount ?? 10));
        writer.Write((uint)recurrence.FirstDayOfWeek);

        DateTime[] modified = recurrence.Exceptions.Select(exception => exception.Start.Date)
            .OrderBy(value => value).ToArray();
        DateTime[] deleted = recurrence.DeletedOccurrenceDates.Select(value => value.Date)
            .Concat(recurrence.Exceptions.Select(exception => exception.OriginalStart.Date))
            .Distinct().OrderBy(value => value).ToArray();
        writer.Write((uint)deleted.Length);
        foreach (DateTime date in deleted) writer.Write(ToMinutes(date));
        writer.Write((uint)modified.Length);
        foreach (DateTime date in modified) writer.Write(ToMinutes(date));
        writer.Write(ToMinutes(recurrence.Start.Date));
        writer.Write(GetWireEndDate(recurrence));
    }

    private static void WriteAppointmentTail(BinaryWriter writer, OutlookRecurrence recurrence, Encoding ansi) {
        OutlookRecurrenceException[] exceptions = recurrence.Exceptions.OrderBy(value => value.Start).ToArray();
        if (exceptions.Length > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(recurrence));
        writer.Write(0x00003006U);
        writer.Write(0x00003009U);
        writer.Write(checked((uint)recurrence.Start.TimeOfDay.TotalMinutes));
        writer.Write(checked((uint)(recurrence.Start.TimeOfDay + recurrence.Duration).TotalMinutes));
        writer.Write((ushort)exceptions.Length);
        var flags = new ushort[exceptions.Length];
        for (int index = 0; index < exceptions.Length; index++) {
            flags[index] = GetOverrideFlags(exceptions[index]);
            WriteException(writer, exceptions[index], flags[index], ansi);
        }
        writer.Write(0U);
        for (int index = 0; index < exceptions.Length; index++)
            WriteExtendedException(writer, exceptions[index], flags[index]);
        writer.Write(0U);
    }

    private static void WriteException(BinaryWriter writer, OutlookRecurrenceException exception, ushort flags,
        Encoding ansi) {
        writer.Write(ToMinutes(exception.Start));
        writer.Write(ToMinutes(exception.End));
        writer.Write(ToMinutes(exception.OriginalStart));
        writer.Write(flags);
        if ((flags & 0x0001) != 0) WriteAnsiWithTwoLengths(writer, exception.Subject!, ansi);
        if ((flags & 0x0002) != 0) writer.Write(unchecked((uint)exception.MeetingType!.Value));
        if ((flags & 0x0004) != 0) writer.Write(unchecked((uint)exception.ReminderDeltaMinutes!.Value));
        if ((flags & 0x0008) != 0) writer.Write(exception.ReminderIsSet == true ? 1U : 0U);
        if ((flags & 0x0010) != 0) WriteAnsiWithTwoLengths(writer, exception.Location!, ansi);
        if ((flags & 0x0020) != 0) writer.Write(unchecked((uint)exception.BusyStatus!.Value));
        if ((flags & 0x0040) != 0) writer.Write(exception.HasAttachments == true ? 1U : 0U);
        if ((flags & 0x0080) != 0) writer.Write(exception.IsAllDay == true ? 1U : 0U);
        if ((flags & 0x0100) != 0) writer.Write(unchecked((uint)exception.AppointmentColor!.Value));
    }

    private static void WriteExtendedException(BinaryWriter writer, OutlookRecurrenceException exception,
        ushort flags) {
        writer.Write(4U);
        writer.Write(0U);
        writer.Write(0U);
        if ((flags & 0x0011) == 0) return;
        writer.Write(ToMinutes(exception.Start));
        writer.Write(ToMinutes(exception.End));
        writer.Write(ToMinutes(exception.OriginalStart));
        if ((flags & 0x0001) != 0) WriteUnicode16(writer, exception.Subject!);
        if ((flags & 0x0010) != 0) WriteUnicode16(writer, exception.Location!);
        writer.Write(0U);
    }

    private static ushort GetOverrideFlags(OutlookRecurrenceException value) {
        int flags = 0;
        if (value.Subject != null) flags |= 0x0001;
        if (value.MeetingType.HasValue) flags |= 0x0002;
        if (value.ReminderDeltaMinutes.HasValue) flags |= 0x0004;
        if (value.ReminderIsSet.HasValue) flags |= 0x0008;
        if (value.Location != null) flags |= 0x0010;
        if (value.BusyStatus.HasValue) flags |= 0x0020;
        if (value.HasAttachments.HasValue) flags |= 0x0040;
        if (value.IsAllDay.HasValue) flags |= 0x0080;
        if (value.AppointmentColor.HasValue) flags |= 0x0100;
        if (value.HasExceptionalBody) flags |= 0x0200;
        return (ushort)flags;
    }

    private static void ValidateForWrite(OutlookRecurrence recurrence, bool appointment) {
        if (recurrence.Start == default) throw new InvalidOperationException("A recurrence requires Start.");
        if (recurrence.Interval <= 0) throw new InvalidOperationException("A recurrence interval must be positive.");
        if (!IsValidFrequencyPattern(recurrence.Frequency, recurrence.PatternKind))
            throw new InvalidOperationException("The recurrence frequency and pattern kind are incompatible.");
        if (recurrence.CalendarType > 2)
            throw new NotSupportedException("Writing non-Gregorian Outlook recurrence calendars is not supported.");
        if (recurrence.Frequency == OutlookRecurrenceFrequency.Daily && recurrence.Interval > 999 ||
            recurrence.Frequency == OutlookRecurrenceFrequency.Weekly && recurrence.Interval > 99 ||
            (recurrence.Frequency == OutlookRecurrenceFrequency.Monthly ||
             recurrence.Frequency == OutlookRecurrenceFrequency.Yearly) && recurrence.Interval > 99)
            throw new ArgumentOutOfRangeException(nameof(recurrence), "The recurrence interval exceeds MAPI limits.");
        if (recurrence.PatternKind == OutlookRecurrencePatternKind.Week && recurrence.DaysOfWeek == OutlookRecurrenceDays.None)
            throw new InvalidOperationException("A weekly recurrence requires weekdays.");
        if (recurrence.PatternKind == OutlookRecurrencePatternKind.MonthDay && !recurrence.DayOfMonth.HasValue)
            throw new InvalidOperationException("A month-day recurrence requires DayOfMonth.");
        if (recurrence.PatternKind == OutlookRecurrencePatternKind.MonthNth &&
            (recurrence.DaysOfWeek == OutlookRecurrenceDays.None || !recurrence.WeekOrdinal.HasValue))
            throw new InvalidOperationException("An ordinal-month recurrence requires weekdays and an ordinal.");
        if (recurrence.RangeKind == OutlookRecurrenceRangeKind.OccurrenceCount && !recurrence.OccurrenceCount.HasValue)
            throw new InvalidOperationException("An occurrence-count recurrence requires OccurrenceCount.");
        if (recurrence.RangeKind == OutlookRecurrenceRangeKind.EndDate && !recurrence.EndDate.HasValue)
            throw new InvalidOperationException("An end-date recurrence requires EndDate.");
        if (!appointment && recurrence.Exceptions.Count != 0)
            throw new NotSupportedException("Task RecurrencePattern values do not contain appointment exceptions.");
        if (appointment) ValidateAppointmentExceptions(recurrence);
        if (recurrence.Duration < TimeSpan.Zero || recurrence.Start.TimeOfDay.TotalMinutes + recurrence.Duration.TotalMinutes > uint.MaxValue)
            throw new ArgumentOutOfRangeException(nameof(recurrence), "The appointment duration cannot be encoded.");
    }

    private static void ValidateAppointmentExceptions(OutlookRecurrence recurrence) {
        if (recurrence.Exceptions.GroupBy(exception => exception.OriginalStart.Date)
                .Any(group => group.Skip(1).Any()))
            throw new InvalidOperationException(
                "Only one Outlook recurrence exception can target each original occurrence date.");
        if (recurrence.Exceptions.GroupBy(exception => exception.Start.Date)
                .Any(group => group.Skip(1).Any()))
            throw new InvalidOperationException(
                "Outlook recurrence exceptions cannot start on the same calendar date.");
        if (recurrence.Exceptions.Any(exception =>
                exception.OriginalStart.TimeOfDay != recurrence.Start.TimeOfDay))
            throw new InvalidOperationException(
                "An Outlook recurrence exception OriginalStart must use the base occurrence time.");

        var originalDates = new HashSet<DateTime>(recurrence.Exceptions.Select(exception =>
            exception.OriginalStart.Date));
        var movedDates = new HashSet<DateTime>(recurrence.Exceptions.Select(exception => exception.Start.Date));
        HashSet<DateTime> baseDates = OutlookRecurrenceExpander.FindBaseOccurrenceDates(recurrence,
            originalDates.Concat(movedDates));
        if (originalDates.Any(date => !baseDates.Contains(date)))
            throw new InvalidOperationException(
                "An Outlook recurrence exception must identify an occurrence in the base series.");

        var deletedDates = new HashSet<DateTime>(recurrence.DeletedOccurrenceDates.Select(value => value.Date));
        if (movedDates.Any(date => baseDates.Contains(date) &&
                !originalDates.Contains(date) && !deletedDates.Contains(date)))
            throw new InvalidOperationException(
                "An Outlook recurrence exception cannot move onto an unmodified base occurrence.");

        Dictionary<DateTime, OutlookRecurrenceNeighborBounds> neighbors =
            OutlookRecurrenceExpander.FindExceptionNeighborBounds(recurrence, originalDates);
        foreach (OutlookRecurrenceException exception in recurrence.Exceptions) {
            if (exception.End < exception.Start)
                throw new InvalidOperationException(
                    "An Outlook recurrence exception cannot end before it starts.");
            OutlookRecurrenceNeighborBounds bounds = neighbors[exception.OriginalStart.Date];
            if ((bounds.PreviousEnd.HasValue && exception.Start < bounds.PreviousEnd.Value) ||
                (bounds.NextStart.HasValue && exception.End > bounds.NextStart.Value))
                throw new InvalidOperationException(
                    "An Outlook recurrence exception must remain between its adjacent effective occurrences without overlap.");
        }
    }

    internal static bool IsValidFrequencyPattern(
        OutlookRecurrenceFrequency frequency,
        OutlookRecurrencePatternKind patternKind) {
        if (frequency == OutlookRecurrenceFrequency.Daily)
            return patternKind == OutlookRecurrencePatternKind.Day;
        if (frequency == OutlookRecurrenceFrequency.Weekly)
            return patternKind == OutlookRecurrencePatternKind.Week;
        if (frequency == OutlookRecurrenceFrequency.Monthly ||
            frequency == OutlookRecurrenceFrequency.Yearly) {
            return patternKind == OutlookRecurrencePatternKind.MonthDay ||
                   patternKind == OutlookRecurrencePatternKind.MonthNth ||
                   patternKind == OutlookRecurrencePatternKind.MonthEnd;
        }
        return false;
    }

    private static uint GetWireEndDate(OutlookRecurrence recurrence) {
        if (recurrence.RangeKind == OutlookRecurrenceRangeKind.NoEnd) return NoEndDate;
        if (recurrence.EndDate.HasValue) return ToMinutes(recurrence.EndDate.Value.Date);
        if (recurrence.RangeKind == OutlookRecurrenceRangeKind.OccurrenceCount) {
            var clone = CloneWithoutExceptions(recurrence);
            OutlookRecurrenceExpansionResult expansion = OutlookRecurrenceExpander.Expand(clone,
                new OutlookRecurrenceExpansionOptions { MaxOccurrences = recurrence.OccurrenceCount!.Value,
                    MaxCandidateDays = 1000000 });
            if (expansion.Occurrences.Count != recurrence.OccurrenceCount.Value)
                throw new InvalidOperationException("The recurrence end date could not be bounded safely.");
            return ToMinutes(expansion.Occurrences[expansion.Occurrences.Count - 1].OriginalStart.Date);
        }
        return ToMinutes(recurrence.Start.Date);
    }

    private static OutlookRecurrence CloneWithoutExceptions(OutlookRecurrence source) {
        return new OutlookRecurrence {
            Frequency = source.Frequency, PatternKind = source.PatternKind, Interval = source.Interval,
            Start = source.Start, Duration = source.Duration, DaysOfWeek = source.DaysOfWeek,
            DayOfMonth = source.DayOfMonth, WeekOrdinal = source.WeekOrdinal,
            FirstDayOfWeek = source.FirstDayOfWeek, RangeKind = source.RangeKind,
            OccurrenceCount = source.OccurrenceCount, EndDate = source.EndDate,
            CalendarType = source.CalendarType, Sliding = source.Sliding, TimeZoneId = source.TimeZoneId
        };
    }

    private static uint ComputeFirstDateTime(OutlookRecurrence recurrence, uint period) {
        DateTime date = recurrence.Start.Date;
        if (recurrence.Frequency == OutlookRecurrenceFrequency.Daily)
            return ToMinutes(date) % period;
        if (recurrence.Frequency == OutlookRecurrenceFrequency.Weekly) {
            int offset = ((int)date.DayOfWeek - (int)recurrence.FirstDayOfWeek + 7) % 7;
            return ToMinutes(date.AddDays(-offset)) % checked(period * 10080U);
        }
        int months = checked((date.Year - 1601) * 12 + date.Month - 1);
        DateTime firstValidMonth = ReferenceDate.AddMonths(months % checked((int)period));
        return ToMinutes(firstValidMonth);
    }

    private static uint EncodePeriod(OutlookRecurrence recurrence) {
        if (recurrence.Frequency == OutlookRecurrenceFrequency.Daily)
            return checked((uint)recurrence.Interval * 1440U);
        if (recurrence.Frequency == OutlookRecurrenceFrequency.Yearly)
            return checked((uint)recurrence.Interval * 12U);
        return checked((uint)recurrence.Interval);
    }

    private static int DecodeInterval(OutlookRecurrenceFrequency frequency, uint period) {
        if (frequency == OutlookRecurrenceFrequency.Daily) {
            if (period == 0 || period % 1440 != 0) throw new InvalidDataException("Invalid daily recurrence period.");
            return checked((int)(period / 1440));
        }
        if (period == 0 || period > int.MaxValue) throw new InvalidDataException("Invalid recurrence period.");
        if (frequency == OutlookRecurrenceFrequency.Yearly && period % 12 != 0)
            throw new InvalidDataException("Invalid yearly recurrence period.");
        return frequency == OutlookRecurrenceFrequency.Yearly ? checked((int)(period / 12)) : (int)period;
    }

    private static OutlookRecurrenceFrequency DecodeFrequency(ushort value) {
        switch (value) {
            case 0x200A: return OutlookRecurrenceFrequency.Daily;
            case 0x200B: return OutlookRecurrenceFrequency.Weekly;
            case 0x200C: return OutlookRecurrenceFrequency.Monthly;
            case 0x200D: return OutlookRecurrenceFrequency.Yearly;
            default: throw new InvalidDataException("Unknown recurrence frequency.");
        }
    }

    private static ushort EncodeFrequency(OutlookRecurrenceFrequency value) =>
        value == OutlookRecurrenceFrequency.Daily ? (ushort)0x200A :
        value == OutlookRecurrenceFrequency.Weekly ? (ushort)0x200B :
        value == OutlookRecurrenceFrequency.Monthly ? (ushort)0x200C : (ushort)0x200D;

    private static OutlookRecurrencePatternKind DecodePatternKind(ushort value) {
        switch (value) {
            case 0x0000: return OutlookRecurrencePatternKind.Day;
            case 0x0001: return OutlookRecurrencePatternKind.Week;
            case 0x0002: return OutlookRecurrencePatternKind.MonthDay;
            case 0x0003: return OutlookRecurrencePatternKind.MonthNth;
            case 0x0004: return OutlookRecurrencePatternKind.MonthEnd;
            default: throw new NotSupportedException("Unsupported recurrence pattern type.");
        }
    }

    private static ushort EncodePatternKind(OutlookRecurrencePatternKind value) =>
        value == OutlookRecurrencePatternKind.Day ? (ushort)0x0000 :
        value == OutlookRecurrencePatternKind.Week ? (ushort)0x0001 :
        value == OutlookRecurrencePatternKind.MonthDay ? (ushort)0x0002 :
        value == OutlookRecurrencePatternKind.MonthNth ? (ushort)0x0003 : (ushort)0x0004;

    private static OutlookRecurrenceRangeKind DecodeRange(uint value) {
        if (value == 0x00002021) return OutlookRecurrenceRangeKind.EndDate;
        if (value == 0x00002022) return OutlookRecurrenceRangeKind.OccurrenceCount;
        if (value == 0x00002023 || value == uint.MaxValue) return OutlookRecurrenceRangeKind.NoEnd;
        throw new InvalidDataException("Unknown recurrence range type.");
    }

    private static uint EncodeRange(OutlookRecurrenceRangeKind value) =>
        value == OutlookRecurrenceRangeKind.EndDate ? 0x00002021U :
        value == OutlookRecurrenceRangeKind.OccurrenceCount ? 0x00002022U : 0x00002023U;

    private static OutlookRecurrenceDays DecodeDays(uint value) {
        if ((value & ~0x7FU) != 0 || (value & 0x7F) == 0) throw new InvalidDataException("Invalid recurrence weekday mask.");
        return (OutlookRecurrenceDays)value;
    }

    private static uint ToMinutes(DateTime value) {
        DateTime local = DateTime.SpecifyKind(value, DateTimeKind.Unspecified);
        long minutes = checked((long)(local - ReferenceDate).TotalMinutes);
        if (minutes < 0 || minutes > uint.MaxValue) throw new ArgumentOutOfRangeException(nameof(value));
        return (uint)minutes;
    }

    private static DateTime FromMinutes(uint value) => ReferenceDate.AddMinutes(value);
    private static int ToInt32(uint value) => unchecked((int)value);

    private static Encoding GetEncoding(int codePage) {
        try {
            return Encoding.GetEncoding(codePage, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
        } catch (ArgumentException) {
            return Encoding.GetEncoding(1252, EncoderFallback.ExceptionFallback, DecoderFallback.ExceptionFallback);
        }
    }

    private static void WriteAnsiWithTwoLengths(BinaryWriter writer, string value, Encoding encoding) {
        byte[] bytes = encoding.GetBytes(value);
        if (bytes.Length >= ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(value));
        writer.Write(checked((ushort)(bytes.Length + 1)));
        writer.Write(checked((ushort)bytes.Length));
        writer.Write(bytes);
    }

    private static void WriteUnicode16(BinaryWriter writer, string value) {
        if (value.Length > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(value));
        writer.Write((ushort)value.Length);
        writer.Write(Encoding.Unicode.GetBytes(value));
    }

    private sealed class Cursor {
        private readonly byte[] _bytes;
        internal Cursor(byte[] bytes) { _bytes = bytes; }
        internal int Offset { get; private set; }
        internal bool AtEnd => Offset == _bytes.Length;
        internal ushort ReadUInt16() { Ensure(2); ushort value = (ushort)(_bytes[Offset] | _bytes[Offset + 1] << 8); Offset += 2; return value; }
        internal uint ReadUInt32() { Ensure(4); uint value = (uint)(_bytes[Offset] | _bytes[Offset + 1] << 8 | _bytes[Offset + 2] << 16 | _bytes[Offset + 3] << 24); Offset += 4; return value; }
        internal uint ReadCount(int maximum, string description) { uint count = ReadUInt32(); if (count > maximum) throw new InvalidDataException("Too many " + description + "."); return count; }
        internal void SkipCountedBlock(string name) { uint length = ReadUInt32(); if (length > int.MaxValue) throw new InvalidDataException(name + " is too large."); Skip((int)length); }
        internal void Skip(int length) { Ensure(length); Offset += length; }
        internal string ReadAnsiWithTwoLengths(Encoding encoding, string name) {
            ushort includingNull = ReadUInt16();
            ushort length = ReadUInt16();
            if (includingNull != length + 1) throw new InvalidDataException("Invalid exception " + name + " length.");
            Ensure(length);
            string value = encoding.GetString(_bytes, Offset, length);
            Offset += length;
            return value;
        }
        internal string ReadUnicode16(string name) {
            ushort characters = ReadUInt16();
            int length = checked(characters * 2);
            Ensure(length);
            string value = Encoding.Unicode.GetString(_bytes, Offset, length);
            Offset += length;
            return value;
        }
        private void Ensure(int length) { if (length < 0 || Offset > _bytes.Length - length) throw new InvalidDataException("The recurrence BLOB is truncated."); }
    }

    private sealed class ParsedPattern {
        internal ParsedPattern(uint modifiedCount, IReadOnlyList<DateTime> modifiedDates) { ModifiedCount = modifiedCount; ModifiedDates = modifiedDates; }
        internal uint ModifiedCount { get; }
        internal IReadOnlyList<DateTime> ModifiedDates { get; }
    }

    private sealed class ParsedException {
        internal ParsedException(OutlookRecurrenceException exception, ushort flags) { Exception = exception; Flags = flags; }
        internal OutlookRecurrenceException Exception { get; }
        internal ushort Flags { get; }
    }
}
