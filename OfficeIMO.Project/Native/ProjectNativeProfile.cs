using OfficeIMO.Core.Internal;

namespace OfficeIMO.Project;

/// <summary>Container identities for independently qualified native generations.</summary>
internal sealed class ProjectNativeProfile {
    internal static readonly ProjectNativeProfile Mpp14 = new ProjectNativeProfile(14);
    internal static readonly ProjectNativeProfile Mpp12 = new ProjectNativeProfile(12);
    internal static readonly ProjectNativeProfile Mpp9 = new ProjectNativeProfile(9);
    internal static readonly ProjectNativeProfile Mpp8 = new ProjectNativeProfile(8);
    internal int Version { get; }
    internal string Header => "Props" + (Version == 8 ? "" : Version.ToString(System.Globalization.CultureInfo.InvariantCulture));
    internal string DataRoot => "   1" + (Version == 8 ? "" : Version.ToString(System.Globalization.CultureInfo.InvariantCulture));
    internal string PresentationRoot => "   2" + (Version == 8 ? "" : Version.ToString(System.Globalization.CultureInfo.InvariantCulture));
    internal string Properties => DataRoot + "/Props";
    internal string Generation => "MPP" + Version;
    internal bool HasExtendedRecords => Version >= 12;
    internal (string Name, ProjectNativeSchema Schema)[] Schemas() => Version switch {
        9 => new[] { ("Task", ProjectNativeSchema9Catalog.Task()), ("Rsc", ProjectNativeSchema9Catalog.Resource()),
            ("Cal", ProjectNativeSchema9Catalog.Calendar()), ("Assn", ProjectNativeSchema9Catalog.Assignment()),
            ("Cons", ProjectNativeSchema9Catalog.Dependency()), ("OutlCode", ProjectNativeSchema9Catalog.OutlineCode()) },
        12 => new[] { ("Task", ProjectNativeSchema12Catalog.Task()), ("Rsc", ProjectNativeSchema12Catalog.Resource()),
            ("Cal", ProjectNativeSchema12Catalog.Calendar()), ("Assn", ProjectNativeSchema12Catalog.Assignment()),
            ("Cons", ProjectNativeSchema12Catalog.Dependency()), ("OutlCode", ProjectNativeSchema12Catalog.OutlineCode()) },
        _ => new[] { ("Task", ProjectNativeSchemaCatalog.Task()), ("Rsc", ProjectNativeSchemaCatalog.Resource()),
            ("Cal", ProjectNativeSchemaCatalog.Calendar()), ("Assn", ProjectNativeSchemaCatalog.Assignment()),
            ("Cons", ProjectNativeSchemaCatalog.Dependency()), ("OutlCode", ProjectNativeSchemaCatalog.OutlineCode()) }
    };
    internal ProjectFileFormat Format(bool template) => Version switch {
        8 => template ? ProjectFileFormat.Mpt8 : ProjectFileFormat.Mpp8,
        9 => template ? ProjectFileFormat.Mpt9 : ProjectFileFormat.Mpp9,
        12 => template ? ProjectFileFormat.Mpt12 : ProjectFileFormat.Mpp12,
        _ => template ? ProjectFileFormat.Mpt14 : ProjectFileFormat.Mpp14
    };
    internal static bool IsTemplate(ProjectFileFormat format) => format == ProjectFileFormat.Mpt8 || format == ProjectFileFormat.Mpt9 || format == ProjectFileFormat.Mpt12 || format == ProjectFileFormat.Mpt14;
    internal static ProjectNativeProfile ForFormat(ProjectFileFormat format) => format switch {
        ProjectFileFormat.Mpp8 or ProjectFileFormat.Mpt8 => Mpp8,
        ProjectFileFormat.Mpp9 or ProjectFileFormat.Mpt9 => Mpp9,
        ProjectFileFormat.Mpp12 or ProjectFileFormat.Mpt12 => Mpp12,
        _ => Mpp14
    };
    private ProjectNativeProfile(int version) { Version = version; }

    internal static ProjectNativeProfile Detect(OfficeCompoundFile file) {
        var matches = new[] { Mpp14, Mpp12, Mpp9, Mpp8 }.Where(p => file.Streams.ContainsKey(p.Header)).ToArray();
        if (matches.Length != 1)
            throw new NotSupportedException("The native container has an unsupported or ambiguous generation header.");
        var profile = matches[0];
        if (!file.Streams.ContainsKey(profile.Properties))
            throw new InvalidDataException("The native generation's project property stream is missing.");
        return profile;
    }
}
