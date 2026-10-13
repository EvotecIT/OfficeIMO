/// <summary>Stages a cut's artifacts beside their final paths and publishes them after production succeeds.</summary>
internal static class MediaArtifacts {
    internal static async Task PublishAsync(string output, string[] names, Func<string, Task> produce) {
        // Sibling staging keeps replacements on the same volume, including still-only recordings.
        string staging = Path.Combine(output, ".format-map-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(staging);
        bool removeStaging = true;
        Exception? failure = null;
        try {
            await produce(staging).ConfigureAwait(false);
            foreach (string name in names) {
                if (!File.Exists(Path.Combine(staging, name))) throw new InvalidOperationException("Missing produced artifact: " + name);
            }
            // Copy all previous outputs before promotion; a locked target may reject a later move.
            string backup = Path.Combine(staging, "previous");
            Directory.CreateDirectory(backup);
            foreach (string name in names) {
                string target = Path.Combine(output, name);
                if (File.Exists(target)) File.Copy(target, Path.Combine(backup, name));
            }
            var published = new List<string>();
            try {
                foreach (string name in names) {
                    File.Move(Path.Combine(staging, name), Path.Combine(output, name), overwrite: true);
                    published.Add(name);
                }
            } catch (Exception publishError) {
                var recoveryErrors = new List<Exception>();
                foreach (string name in published.AsEnumerable().Reverse()) {
                    try {
                        string previous = Path.Combine(backup, name);
                        string target = Path.Combine(output, name);
                        if (File.Exists(previous)) File.Move(previous, target, overwrite: true);
                        else File.Delete(target);
                    } catch (Exception recoveryError) { recoveryErrors.Add(recoveryError); }
                }
                if (recoveryErrors.Count > 0) {
                    removeStaging = false;
                    throw new AggregateException($"Publishing and recovery failed; recovery files remain at '{staging}'.", [publishError, .. recoveryErrors]);
                }
                throw;
            }
            foreach (string name in names) {
                string target = Path.Combine(output, name);
                Console.WriteLine($"  {name,-38} {new FileInfo(target).Length / 1024.0 / 1024.0,7:0.00} MB");
            }
        } catch (Exception error) {
            failure = error;
            throw;
        } finally {
            if (removeStaging) {
                try {
                    // Backups inherit attributes. Clear read-only only on disposable staging copies.
                    foreach (string file in Directory.EnumerateFiles(staging, "*", SearchOption.AllDirectories)) {
                        FileAttributes attributes = File.GetAttributes(file);
                        if ((attributes & FileAttributes.ReadOnly) != 0) File.SetAttributes(file, attributes & ~FileAttributes.ReadOnly);
                    }
                    Directory.Delete(staging, recursive: true);
                } catch (Exception cleanupError) when (failure is not null) {
                    Console.Error.WriteLine($"Cleanup failed; staging files remain at '{staging}': {cleanupError.Message}");
                }
            }
        }
    }
}
