/// <summary>Stages a cut's artifacts beside their final paths and publishes them after production succeeds.</summary>
internal static class MediaArtifacts {
    internal static async Task PublishAsync(string output, string[] names, Func<string, Task> produce) {
        // Sibling staging keeps replacements on the same volume, including still-only recordings.
        string staging = Path.Combine(output, ".format-map-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(staging);
        try {
            await produce(staging).ConfigureAwait(false);
            foreach (string name in names) {
                if (!File.Exists(Path.Combine(staging, name))) throw new InvalidOperationException("Missing produced artifact: " + name);
            }
            foreach (string name in names) {
                string target = Path.Combine(output, name);
                File.Move(Path.Combine(staging, name), target, overwrite: true);
                Console.WriteLine($"  {name,-38} {new FileInfo(target).Length / 1024.0 / 1024.0,7:0.00} MB");
            }
        } finally {
            Directory.Delete(staging, recursive: true);
        }
    }
}
