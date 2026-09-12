namespace OfficeIMO.Project;

public sealed partial class ProjectDocument {
    private ProjectFileFormat _associatedFormat;

    /// <summary>Loads a document template into an unassociated editable project; saving cannot overwrite the template implicitly.</summary>
    /// <remarks>Template objects, identities, stored dates, and inert content are retained. This does not instantiate Global.mpt or execute macros.</remarks>
    public static ProjectDocument CreateFromTemplate(string path, ProjectLoadOptions? options = null, CancellationToken cancellationToken = default) {
        if (options?.PersistenceMode == DocumentPersistenceMode.SaveOnDispose || options?.AccessMode == DocumentAccessMode.ReadOnly)
            throw new ArgumentException("Template authoring requires explicit persistence and read/write access.", nameof(options));
        var document = Load(path, options, cancellationToken);
        document._path = null; document._associatedStream = null; document._associatedFormat = document.NativeInfo?.Profile.Format(false) ?? ProjectFileFormat.Mpp14;
        document._savedRevision = -1;
        return document;
    }

    /// <summary>Loads caller-owned template bytes into an unassociated editable project. The input is never rebound as the save destination.</summary>
    public static ProjectDocument CreateFromTemplate(Stream stream, ProjectLoadOptions? options = null, CancellationToken cancellationToken = default) {
        if (options?.PersistenceMode == DocumentPersistenceMode.SaveOnDispose || options?.AccessMode == DocumentAccessMode.ReadOnly)
            throw new ArgumentException("Template authoring requires explicit persistence and read/write access.", nameof(options));
        var document = Load(stream, options, cancellationToken);
        document._path = null; document._associatedStream = null; document._associatedFormat = document.NativeInfo?.Profile.Format(false) ?? ProjectFileFormat.Mpp14;
        document._savedRevision = -1;
        return document;
    }

    private ProjectFileFormat ResolveFormat(ProjectSaveOptions options, string? path = null) {
        options.Validate();
        ProjectFileFormat requested = options.Format;
        if (path != null) {
            var inferred = FormatFromPath(path);
            if (requested != ProjectFileFormat.Automatic && requested != inferred &&
                !(inferred == ProjectFileFormat.Mpp14 && (requested == ProjectFileFormat.Mpp12 || requested == ProjectFileFormat.Mpp9 || requested == ProjectFileFormat.Mpp8)) &&
                !(inferred == ProjectFileFormat.Mpt14 && (requested == ProjectFileFormat.Mpt12 || requested == ProjectFileFormat.Mpt9 || requested == ProjectFileFormat.Mpt8)))
                throw new ArgumentException("The requested Project format does not match the destination extension.", nameof(options));
            if (requested != ProjectFileFormat.Automatic) return requested;
            if (IsNativeFormat(inferred)) {
                var profile = IsNativeFormat(_associatedFormat)
                    ? ProjectNativeProfile.ForFormat(_associatedFormat) : NativeInfo?.Profile;
                if (profile != null) return profile.Format(inferred == ProjectFileFormat.Mpt14);
            }
            return inferred;
        }
        if (requested != ProjectFileFormat.Automatic) return requested;
        if (_associatedFormat != ProjectFileFormat.Automatic) return _associatedFormat;
        return MpxSource != null ? ProjectFileFormat.Mpx4 : NativeInfo == null ? ProjectFileFormat.Xml : NativeInfo.Profile.Format(NativeInfo.IsTemplate);
    }
    private static bool IsNativeFormat(ProjectFileFormat format) => format != ProjectFileFormat.Automatic && format != ProjectFileFormat.Xml && format != ProjectFileFormat.Mpx4;
    private static ProjectFileFormat FormatFromPath(string path) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("A Project output path is required.", nameof(path));
        RejectGlobalTemplate(path);
        return System.IO.Path.GetExtension(path).ToLowerInvariant() switch {
            ".xml" => ProjectFileFormat.Xml, ".mpp" => ProjectFileFormat.Mpp14, ".mpt" => ProjectFileFormat.Mpt14,
            ".mpx" => ProjectFileFormat.Mpx4,
            _ => throw new NotSupportedException("Qualified Project output extensions are .xml, .mpp, .mpt, and .mpx.") };
    }
    private static void RejectGlobalTemplate(string path) {
        if (string.Equals(System.IO.Path.GetFileName(path), "Global.mpt", StringComparison.OrdinalIgnoreCase))
            throw new NotSupportedException("Global.mpt is an application-wide store, not a qualified document template.");
    }
    private static ProjectSaveOptions WithFormat(ProjectSaveOptions options, ProjectFileFormat format, bool? preserveBytes = null) => new ProjectSaveOptions {
        Format = format, MpxEncoding = options.MpxEncoding, MpxSeparator = options.MpxSeparator, Indent = options.Indent, LossPolicy = options.LossPolicy, FileConflictPolicy = options.FileConflictPolicy,
        MaxOutputBytes = options.MaxOutputBytes, PreserveUnchangedBytes = preserveBytes ?? options.PreserveUnchangedBytes
    };

    private bool CanRetainNative(ProjectSaveOptions options, ProjectFileFormat format) => NativeSource != null && NativeSource.ModelRevision == Revision && !_batchChanged
        && options.PreserveUnchangedBytes && format != ProjectFileFormat.Xml && format == NativeInfo!.Profile.Format(NativeInfo.IsTemplate);

    private ProjectReport AssessFormat(ProjectSaveOptions options, ProjectFileFormat format, CancellationToken token, bool includeNativePlan = true) {
        token.ThrowIfCancellationRequested();
        if (CanRetainNative(options, format)) {
            var retained = NativeSource!.UnrepresentedValues.ToList();
            AddRetainedOutputLimit(retained, NativeSource.Bytes, options);
            return new ProjectReport(Revision, retained);
        }
        if (format == ProjectFileFormat.Mpx4 && ProjectMpxWriter.CanRetain(this, options)) return ProjectMpxWriter.Plan(this, options, false, token).Report;
        var diagnostics = Validate(token).Diagnostics.ToList();
        if (format == ProjectFileFormat.Xml) {
            foreach (var task in AllTasks) {
                token.ThrowIfCancellationRequested();
                if (!ProjectXmlValue.TryTaskDurationFormat(task, out _))
                    diagnostics.Add(new ProjectDiagnostic("PROJECT_XML_DURATION_FORMAT", ProjectDiagnosticSeverity.Error,
                        "XML task duration, actual duration, and remaining duration share one format. Use the same unit and flags.", "/Task[UID=" + task.Uid + "]"));
            }
        }
        if (format == ProjectFileFormat.Xml && NativeSource == null && MpxSource == null
            && ProjectXmlCodec.RetainedBytes(this, options) is byte[] retainedXml)
            AddRetainedOutputLimit(diagnostics, retainedXml, options);
        if (format == ProjectFileFormat.Mpx4) {
            diagnostics.RemoveAll(d => d.Code == "PROJECT_OPAQUE_REFERENCES");
            if (includeNativePlan && !diagnostics.Any(d => d.Severity == ProjectDiagnosticSeverity.Error))
                diagnostics.AddRange(ProjectMpxWriter.Plan(this, WithFormat(options, format), false, token).Report.Diagnostics);
        } else if (IsNativeFormat(format)) {
            diagnostics.RemoveAll(d => d.Code == "PROJECT_OPAQUE_REFERENCES");
            if (MpxSource != null) diagnostics.AddRange(MpxSource.Unmodeled.Select(d => new ProjectDiagnostic("PROJECT_MPX_CONVERSION_LOSS", ProjectDiagnosticSeverity.Warning, d.Message + " This source content is omitted during conversion.", d.Location, true)));
            if (includeNativePlan && !diagnostics.Any(d => d.Severity == ProjectDiagnosticSeverity.Error))
                diagnostics.AddRange(ProjectNativeWriter.Plan(this, WithFormat(options, format), false, token).Report.Diagnostics);
        } else if (NativeSource != null) {
            if (Settings.CurrencyCode == null) diagnostics.Add(new ProjectDiagnostic("PROJECT_CURRENCY_CODE", ProjectDiagnosticSeverity.Error,
                "Set a currency code before converting this native project to MSPDI.", "/Project/CurrencyCode"));
            void Loss(string code, string message, string location = "/") => diagnostics.Add(new ProjectDiagnostic(code, ProjectDiagnosticSeverity.Warning, message, location, true));
            Loss("PROJECT_NATIVE_PRESENTATION_LOSS", "Native views, filters, groups, drawings, report definitions, and other unmodeled records are omitted from XML output.");
            Loss("PROJECT_NATIVE_CURVE_LOSS", "Native timephased work/cost curves and rate tables are not decoded into XML. Modeled scalar totals are retained.");
            Loss("PROJECT_NATIVE_CUSTOM_METADATA_LOSS", "Unmodeled custom-field formulas, lookups, graphical indicators, and enterprise metadata are omitted. Modeled aliases and scalar values are retained.");
            foreach (var diagnostic in ReadDiagnostics.Where(d => d.Code == "PROJECT_NATIVE_RTF_NOTES" || d.Code == "PROJECT_NATIVE_CALENDAR_METADATA"))
                Loss(diagnostic.Code + "_LOSS", diagnostic.Message, diagnostic.Location);
            if (NativeInfo!.HasMacroStorage) Loss("PROJECT_NATIVE_MACRO_LOSS", "Inert native VBA storage is omitted from XML.");
            if (NativeInfo.HasEmbeddedContent) Loss("PROJECT_NATIVE_EMBEDDED_LOSS", "Inert embedded native content is omitted from XML.");
            if (NativeInfo.HasSignatureStorage) Loss("PROJECT_NATIVE_SIGNATURE_LOSS", "Native signatures do not apply to the converted XML and are omitted.");
        }
        if (format == ProjectFileFormat.Xml && MpxSource != null) {
            if (Settings.CurrencyCode == null) diagnostics.Add(new ProjectDiagnostic("PROJECT_CURRENCY_CODE", ProjectDiagnosticSeverity.Error,
                "Set a currency code before converting MPX to MSPDI; an MPX symbol does not identify a currency unambiguously.", "/Project/CurrencyCode"));
            diagnostics.AddRange(MpxSource.Unmodeled.Select(d => new ProjectDiagnostic("PROJECT_MPX_CONVERSION_LOSS", ProjectDiagnosticSeverity.Warning, d.Message + " This source content is omitted during conversion.", d.Location, true)));
        }
        if (format != ProjectFileFormat.Mpx4 && MpxSource?.Comments.Count > 0) diagnostics.Add(new ProjectDiagnostic("PROJECT_MPX_COMMENT_LOSS", ProjectDiagnosticSeverity.Warning,
            "MPX comment records have no model mapping and are omitted during conversion.", "/", true));
        return new ProjectReport(Revision, diagnostics);
    }
    private static void AddRetainedOutputLimit(List<ProjectDiagnostic> diagnostics, byte[] bytes, ProjectSaveOptions options) {
        if (bytes.LongLength > options.MaxOutputBytes) diagnostics.Add(new ProjectDiagnostic("PROJECT_OUTPUT_LIMIT",
            ProjectDiagnosticSeverity.Error, "Retained output exceeds MaxOutputBytes.", "/"));
    }
}
