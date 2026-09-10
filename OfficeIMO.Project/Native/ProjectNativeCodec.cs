using OfficeIMO.Core.Internal;

namespace OfficeIMO.Project;

internal static partial class ProjectNativeCodec {
    internal static bool IsCompound(byte[] bytes) => bytes.Length >= 8 && bytes[0] == 0xd0 && bytes[1] == 0xcf && bytes[2] == 0x11 && bytes[3] == 0xe0
        && bytes[4] == 0xa1 && bytes[5] == 0xb1 && bytes[6] == 0x1a && bytes[7] == 0xe1;

    internal static ProjectDocument Read(byte[] bytes, ProjectLoadOptions options, CancellationToken token) {
        token.ThrowIfCancellationRequested();
        var limits = new OfficeCompoundReadOptions(options.MaxCompoundEntries, options.MaxCompoundStreams, options.MaxInputBytes, options.MaxInputBytes);
        if (!OfficeCompoundFileReader.TryRead(bytes, limits, token, out var file, out var error) || file == null)
            throw new InvalidDataException("MPP compound container: " + error);
        token.ThrowIfCancellationRequested();
        var profile = ProjectNativeProfile.Detect(file);
        var headerBytes = file.Streams[profile.Header];
        var propertyBytes = file.Streams[profile.Properties];
        var header = ProjectNativeProperties.Read(headerBytes, token);
        foreach (uint key in new[] { 0x35400000u, 0x35400001u }) {
            if (header.TryGetValue(key, out var protection) && protection.Copy().Any(b => b != 0))
                throw new NotSupportedException("Protected MPP input is not qualified. Remove protection in the producing application before inspection.");
        }
        var properties = ProjectNativeProperties.Read(propertyBytes, token, profile == ProjectNativeProfile.Mpp8);
        // Current Project writes a value of 14 when exporting a Props12/112 container.
        // Container generation comes from its stream identities, not this producer property alone.
        bool hasGeneration = properties.TryGetValue(0x35400016, out var generation);
        if (!hasGeneration && profile == ProjectNativeProfile.Mpp14 || hasGeneration &&
            generation.Int32() != profile.Version && !(profile == ProjectNativeProfile.Mpp12 && generation.Int32() == 14))
            throw new NotSupportedException("The native project property generation does not match its container.");
        var document = ProjectDocument.CreateForRead();
        document.ReadDiagnosticLimit = options.MaxDiagnostics;
        try {
            document.NativeSource = new ProjectNativeSource(bytes, new ProjectNativeInfo(header.TryGetValue(0x35400010, out var producer) ? producer.Unicode() : null, file));
            ReadSettings(document, properties);
            ReadMetadata(document, file, token);
            ReadCustomAliases(document, file, profile, token);
            var calendarTable = Table(file, profile, properties, "Cal", 0x16, options, token);
            ReadCalendars(document, calendarTable, properties, options, token);
            ReadTasks(document, Table(file, profile, properties, "Task", 0x14, options, token), options, token);
            ReadResources(document, Table(file, profile, properties, "Rsc", 0x15, options, token), options, token);
            ReadAssignments(document, Table(file, profile, properties, "Assn", 0x17, options, token), options, token);
            ReadDependencies(document, Table(file, profile, properties, "Cons", 0x18, options, token), token);
            if ((long)document.TaskIndex.Count + document.ResourceIndex.Count + document.CalendarIndex.Count + document.AssignmentIndex.Count > options.MaxEntities)
                throw new InvalidDataException("Native combined entity budget exceeded.");
            Warn(document, "PROJECT_NATIVE_OPAQUE", "Source streams and unmodeled records are retained. Unchanged native save preserves the whole file; native edits and conversion require a separately qualified writer.", "/");
            Warn(document, "PROJECT_NATIVE_TIMEPHASED_PRESERVED", "Native work/cost curves, rate tables, custom-field formulas/lookups, and presentation records remain in source streams. They are not expanded into typed timephased values or evaluated.", "/");
            document.FinishRead(options);
            document.NativeSource.Snapshot = ProjectModelSnapshot.Capture(document, token);
            document.NativeSource.ModelRevision = document.Revision;
            return document;
        } catch { document.DisposeFailedRead(); throw; }
    }
    private static ProjectNativeTable Table(OfficeCompoundFile file, ProjectNativeProfile profile, Dictionary<uint, ProjectNativeValue> properties, string name, uint id,
        ProjectLoadOptions options, CancellationToken token) {
        if (profile == ProjectNativeProfile.Mpp8) {
            if (!properties.TryGetValue(0x02000000 | (id - 0x13), out var descriptor)) throw new InvalidDataException("Missing legacy field table: " + name);
            return new ProjectNativeTable(file, "TBknd" + name, descriptor, checked(options.MaxEntities + 16), token);
        }
        if (!properties.TryGetValue(0x03000000 | id, out var first)) throw new InvalidDataException("Missing field table: " + name);
        return new ProjectNativeTable(file, "TBknd" + name, first, properties.TryGetValue(0x00020000 | id, out var second) ? second : (ProjectNativeValue?)null,
            checked(options.MaxEntities + 16), token, profile);
    }
    private static void Warn(ProjectDocument document, string code, string message, string location) => document.AddReadDiagnostic(new ProjectDiagnostic(code, ProjectDiagnosticSeverity.Warning, message, location));
    private static void CheckEntityBudget(ProjectDocument document, ProjectLoadOptions options) {
        if ((long)document.TaskIndex.Count + document.ResourceIndex.Count + document.CalendarIndex.Count + document.AssignmentIndex.Count >= options.MaxEntities)
            throw new InvalidDataException("Native combined entity budget exceeded.");
    }
    private static ProjectWork? Work(ProjectNativeRecord record, uint id) {
        decimal? value = record.Number(id); return value.HasValue ? new ProjectWork(value.Value / 1000m) : (ProjectWork?)null;
    }
    private static ProjectDuration? Duration(ProjectNativeRecord record, uint id, uint formatId, ProjectDocument document) {
        int? value = record.Integer(id);
        if (!value.HasValue) return null;
        int format = record.Integer(formatId) ?? 7;
        bool estimated = (format & 32) != 0;
        format &= ~32;
        if (format == 21) format = 7;
        if (format < 3 || format > 12) throw new NotSupportedException("Unqualified native duration unit: " + format);
        var unit = (ProjectDurationUnit)((format - 3) / 2);
        bool elapsed = format % 2 == 0;
        return new ProjectDuration(value.Value / 10m / ProjectXmlValue.MinutesPerUnit(unit, elapsed, document), unit, elapsed, estimated);
    }
    private static void ReadSettings(ProjectDocument document, Dictionary<uint, ProjectNativeValue> values) {
        ProjectNativeValue? Get(uint id) => values.TryGetValue(id, out var value) ? value : (ProjectNativeValue?)null;
        document.Name = Get(0x02400008)?.Unicode();
        if (Get(0x02400029) is ProjectNativeValue identity)
            document.Guid = document.NativeInfo!.Profile.Version <= 9 ? Guid.Parse(identity.Unicode()) : new Guid(identity.Copy());
        document.Settings.StartDate = Get(0x02400002)?.Date(); document.Settings.FinishDate = Get(0x02400003)?.Date();
        document.Settings.ScheduleFromStart = Get(0x02400004)?.UInt16() == 1;
        document.Settings.MinutesPerDay = Get(0x0240001d)?.Int32(); document.Settings.MinutesPerWeek = Get(0x0240001e)?.Int32();
        document.Settings.DaysPerMonth = Get(0x0240138f)?.UInt16();
        document.Settings.CurrencyCode = Get(0x024013bb)?.Unicode(); document.Settings.CurrencySymbol = Get(0x02400010)?.Unicode();
        document.Settings.CurrencyDigits = Get(0x02400012)?.UInt16();
        document.Settings.DefaultStartTime = TimeSpan.FromMinutes((Get(0x0240001c)?.UInt16() ?? 4800) / 10d);
        document.Settings.DefaultFinishTime = TimeSpan.FromMinutes((Get(0x02400021)?.UInt16() ?? 10200) / 10d);
        document.Settings.StatusDate = Get(0x0240003e)?.Date();
    }
}
