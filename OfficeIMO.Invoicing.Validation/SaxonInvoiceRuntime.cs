namespace OfficeIMO.Invoicing.Validation;

/// <summary>Captures the complete pinned JAR class path before any rule subprocess starts.</summary>
internal sealed class SaxonInvoiceRuntime {
    private static readonly (string Path, string Hash, int MaximumBytes)[] Files = {
        ("saxon-he-12.10.jar", SaxonInvoiceRulesRunner.JarSha256, 8 * 1024 * 1024),
        ("lib/xmlresolver-5.3.3.jar", "1FE4D5B92F708DCDB82DBCE12919E0171E6B5CA62C6DCA6220483625098FEB5F", 2 * 1024 * 1024),
        ("lib/xmlresolver-5.3.3-data.jar", "B0C487AD2F3E558BE8D829C916D2458D10ACA6A5BAFA7A4D0524B70845E48A5C", 2 * 1024 * 1024),
        ("lib/jline-2.14.6.jar", "97D1ACAAC82409BE42E622D7A54D3AE9D08517E8AEFDEA3D2BA9791150C2F02D", 1024 * 1024)
    };
    private readonly byte[][] _snapshots;

    private SaxonInvoiceRuntime(byte[][] snapshots) => _snapshots = snapshots;

    internal static string Identity => "SaxonJ-HE 12.10; " + string.Join("; ", Files.Select(file => file.Path + " SHA256=" + file.Hash));

    internal static SaxonInvoiceRuntime Load(string mainJar) {
        string root = Path.GetDirectoryName(mainJar)!;
        var snapshots = new byte[Files.Length][];
        for (int index = 0; index < Files.Length; index++) {
            var file = Files[index];
            snapshots[index] = InvoiceRuleBundle.ReadPinned(index == 0 ? mainJar : Path.Combine(root, file.Path), file.Hash, file.MaximumBytes);
        }
        return new SaxonInvoiceRuntime(snapshots);
    }

    internal async Task<string> WriteAsync(string workingDirectory, CancellationToken cancellationToken) {
        string root = Path.Combine(workingDirectory, "runtime");
        Directory.CreateDirectory(Path.Combine(root, "lib"));
        for (int index = 0; index < Files.Length; index++)
            await File.WriteAllBytesAsync(Path.Combine(root, Files[index].Path), _snapshots[index], cancellationToken).ConfigureAwait(false);
        // Only verified canonical files are copied. Manifest aliases and loose classes from
        // the configured installation cannot join the subprocess class path.
        return Path.Combine(root, Files[0].Path);
    }
}
