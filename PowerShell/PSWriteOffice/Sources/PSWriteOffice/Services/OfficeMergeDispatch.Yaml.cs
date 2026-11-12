using OfficeIMO.Yaml;
using PSWriteOffice.Foundation;

namespace PSWriteOffice.Services;

internal static partial class OfficeMergeDispatch {
    /// <summary>
    /// Overlays each YAML input onto the target in order, the way Helm combines values files, and returns one native
    /// report per overlay. The target is the live document, or the first input; only its edited text changes on save.
    /// </summary>
    private static YamlMergeReport[] MergeYaml(IReadOnlyList<string> paths, YamlMergeOptions options, YamlDocument? document, string? output, bool noClobber) {
        // Load every input before the live target changes, so a malformed later file cannot leave a partial merge.
        var inputs = new List<YamlDocument>(paths.Count);
        foreach (string path in paths) inputs.Add(YamlDocument.Load(path));
        YamlDocument target = document ?? inputs[0];
        var reports = new List<YamlMergeReport>(inputs.Count);
        for (int index = document == null ? 1 : 0; index < inputs.Count; index++) reports.Add(target.Merge(inputs[index], options));
        if (output != null) OfficeOutput.Write(output, stream => target.Save(stream), noClobber);
        return reports.ToArray();
    }
}
