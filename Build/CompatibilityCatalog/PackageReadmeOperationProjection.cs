using System.Text;
using OfficeIMO;

internal static class PackageReadmeOperationProjection {
    private const string StartMarker = "<!-- officeimo-operation-catalog:start -->";
    private const string EndMarker = "<!-- officeimo-operation-catalog:end -->";

    internal static IReadOnlyDictionary<string, string> Create(
        string repositoryRoot,
        IReadOnlyList<OfficeOperationCapability> capabilities) {
        string root = Path.GetFullPath(repositoryRoot);
        var outputs = new SortedDictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (IGrouping<string, OfficeOperationCapability> package in capabilities
                     .GroupBy(static row => row.PackageId, StringComparer.Ordinal)
                     .OrderBy(static group => group.Key, StringComparer.Ordinal)) {
            string readmePath = Path.Combine(root, package.Key, "README.md");
            if (!File.Exists(readmePath)) continue;
            string current = Normalize(File.ReadAllText(readmePath));
            outputs[readmePath] = ReplaceBlock(current, Render(package.Key, package.ToArray()));
        }
        return outputs;
    }

    private static string Render(string packageId, IReadOnlyList<OfficeOperationCapability> rows) {
        var output = new StringBuilder();
        output.AppendLine(StartMarker)
            .AppendLine("## Generated capability summary")
            .AppendLine()
            .AppendLine("This table is generated from the package-neutral OfficeIMO operation catalog. The detailed source contracts remain authoritative for feature-level behavior and limitations.")
            .AppendLine()
            .AppendLine("| Operation | Supported | Partial | Preserved | Rejected | Unsupported | Not applicable |")
            .AppendLine("| --- | ---: | ---: | ---: | ---: | ---: | ---: |");
        foreach (IGrouping<OfficeOperationKind, OfficeOperationCapability> operation in rows
                     .GroupBy(static row => row.Operation)
                     .OrderBy(static group => group.Key)) {
            output.Append("| ").Append(operation.Key);
            foreach (OfficeOperationSupportState state in Enum.GetValues<OfficeOperationSupportState>()) {
                output.Append(" | ").Append(operation.Count(row => row.State == state));
            }
            output.AppendLine(" |");
        }
        output.AppendLine()
            .Append("The complete rows for `").Append(packageId)
            .AppendLine("` are published in the [generated operation contract](https://github.com/EvotecIT/OfficeIMO/blob/master/Docs/Compatibility/generated/package-operations.md).")
            .AppendLine(EndMarker);
        return output.ToString();
    }

    private static string ReplaceBlock(string content, string block) {
        int start = content.IndexOf(StartMarker, StringComparison.Ordinal);
        int end = content.IndexOf(EndMarker, StringComparison.Ordinal);
        if ((start < 0) != (end < 0) || start >= 0 && end < start) {
            throw new InvalidDataException("Package README contains an incomplete generated operation block.");
        }

        string withoutBlock;
        if (start >= 0) {
            int after = end + EndMarker.Length;
            withoutBlock = content.Remove(start, after - start).TrimEnd('\n');
        } else {
            withoutBlock = content.TrimEnd('\n');
        }
        return withoutBlock + "\n\n" + block.TrimEnd('\n') + "\n";
    }

    private static string Normalize(string value) => value
        .Replace("\r\n", "\n")
        .Replace("\r", "\n");
}
