using OfficeIMO.Toml;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services;

internal static partial class OfficeMergeDispatch {
    /// <summary>
    /// Overlays each TOML input onto the target in order, as layered configuration files combine, and returns one native
    /// report per overlay. The target is the live document, or the first input; only its edited text changes on save.
    /// </summary>
    private static TomlMergeReport[] MergeToml(IReadOnlyList<string> paths, TomlMergeOptions options, TomlDocument? document, string? output, bool noClobber) {
        // Load every input before the live target changes, so a malformed later file cannot leave a partial merge.
        var inputs = new List<TomlDocument>(paths.Count);
        foreach (string path in paths) inputs.Add(TomlDocument.Load(path));
        TomlDocument target = document ?? inputs[0];
        var reports = new List<TomlMergeReport>(inputs.Count);
        for (int index = document == null ? 1 : 0; index < inputs.Count; index++) reports.Add(target.Merge(inputs[index], options));
        if (output != null) OfficeOutput.Write(output, stream => target.Save(stream), noClobber);
        return reports.ToArray();
    }
}
