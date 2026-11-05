using OfficeIMO.Core.Internal;
using OfficeIMO.Internal;
using OfficeIMO.OneNote;

namespace OfficeIMO.Workflows;

public static partial class OfficeWorkflow {
    // TOC streams intentionally expose metadata only. Capture the native reader's referenced local files before projection.
    private sealed class OneNoteNotebookCapture : IDisposable {
        private readonly string temporaryRoot = OfficeTemporaryDirectory.Create("officeimo-notebook-");
        private readonly OneNoteNotebookReaderOptions options;
        private readonly long maximumBytes;
        private long bytes;
        private int entries;
        private int fileEntries = 1;
        internal string TableOfContentsPath { get; private set; } = "";
        internal OneNoteNotebookGuard Guard { get; }

        private OneNoteNotebookCapture(string source, OneNoteNotebookReaderOptions options, long maximumBytes) {
            this.options = options;
            this.maximumBytes = Math.Min(maximumBytes, options.MaxPackageExpandedBytes);
            Guard = new OneNoteNotebookGuard(Path.GetDirectoryName(source)!, this.maximumBytes);
        }

        internal static async Task<OneNoteNotebookCapture> CreateAsync(Stream input, string source,
            OneNoteNotebookReaderOptions options, long maximumBytes, CancellationToken token) {
            if (options.MaxNotebookEntries < 1 || options.MaxPackageEntries < 1 || options.MaxPackageEntryBytes < 1 ||
                options.MaxPackageExpandedBytes < 1 || options.MaxSectionGroupDepth < 0)
                throw new ArgumentException("OneNote notebook limits must be positive; group depth can be zero.", nameof(options));
            source = OfficePathIdentity.ResolvePhysicalPath(source);
            var capture = new OneNoteNotebookCapture(source, options, maximumBytes);
            try {
                // Retain the real folder name: it is the native notebook's display name.
                string name = Path.GetFileName(Path.GetDirectoryName(source)) ?? string.Empty;
                string directory = Path.Combine(capture.temporaryRoot, string.IsNullOrEmpty(name) ? "notebook" : name);
                Directory.CreateDirectory(directory);
                capture.TableOfContentsPath = Path.Combine(directory, Path.GetFileName(source));
                input.Position = 0;
                using (var output = OfficeTemporaryFile.CreateAtPath(capture.TableOfContentsPath, 81920, FileOptions.Asynchronous))
                    await input.CopyToAsync(output, 81920, token).ConfigureAwait(false);
                capture.bytes = input.Length;
                if (capture.bytes > capture.maximumBytes || capture.bytes > options.MaxPackageEntryBytes)
                    throw new InvalidDataException("The OneNote notebook exceeds its capture byte limit.");
                string fingerprint = await FingerprintAsync(capture.TableOfContentsPath, capture.maximumBytes, token).ConfigureAwait(false);
                capture.Guard.Files.Add(source, fingerprint);
                await capture.CaptureTocAsync(source, capture.TableOfContentsPath, 0, token).ConfigureAwait(false);
                await capture.Guard.VerifyAsync(token).ConfigureAwait(false);
                return capture;
            } catch { capture.Dispose(); throw; }
        }

        private async Task CaptureTocAsync(string source, string target, int depth, CancellationToken token) {
            token.ThrowIfCancellationRequested();
            if (depth > options.MaxSectionGroupDepth) throw new InvalidDataException("The notebook exceeds its section-group depth limit.");
            OneNoteNotebookReaderOptions metadataOptions = OfficeWorkflowRouteSettings.Snapshot(options)!;
            metadataOptions.LoadSectionContent = false;
            using var stream = new FileStream(target, FileMode.Open, FileAccess.Read, FileShare.Read);
            OneNoteNotebook toc = OneNoteNotebookReader.Read(stream, Path.GetFileName(source), metadataOptions);
            string sourceDirectory = Path.GetDirectoryName(source)!, targetDirectory = Path.GetDirectoryName(target)!;
            foreach (OneNoteSection section in toc.Sections) {
                CountEntry();
                if (options.LoadSectionContent)
                    await CaptureFileAsync(Path.Combine(sourceDirectory, section.SourcePath!),
                        Path.Combine(targetDirectory, section.SourcePath!), token).ConfigureAwait(false);
            }
            foreach (OneNoteSectionGroup group in toc.SectionGroups) {
                CountEntry();
                if (!options.RecurseSectionGroups) continue;
                string childDirectory = Path.Combine(sourceDirectory, group.RelativePath!);
                bool exists = IsRegularEntry(childDirectory, directory: true);
                string? childToc = FindToc(childDirectory);
                Guard.Groups.Add(childDirectory, (exists, childToc));
                if (exists) Directory.CreateDirectory(Path.Combine(targetDirectory, group.RelativePath!));
                if (childToc == null) continue;
                string childTarget = Path.Combine(targetDirectory, group.RelativePath!, Path.GetFileName(childToc));
                await CaptureFileAsync(childToc, childTarget, token).ConfigureAwait(false);
                await CaptureTocAsync(childToc, childTarget, depth + 1, token).ConfigureAwait(false);
            }
        }

        private void CountEntry() {
            if (++entries > options.MaxNotebookEntries)
                throw new InvalidDataException("The notebook exceeds its capture entry limit.");
        }

        private async Task CaptureFileAsync(string source, string target, CancellationToken token) {
            token.ThrowIfCancellationRequested();
            if (!IsRegularEntry(source, directory: false)) { Guard.Files.Add(source, null); return; }
            if (++fileEntries > options.MaxPackageEntries) throw new InvalidDataException("The notebook exceeds its capture file limit.");
            long limit = Math.Min(maximumBytes - bytes, Math.Min(options.MaxPackageEntryBytes, options.OneNoteOptions.MaxInputBytes ?? maximumBytes));
            using var captured = await OfficeStreamFileSnapshot.CaptureAsync(_ => Task.FromResult<Stream>(
                OfficePathIdentity.OpenRegularFileForRead(source, Guard.SourceRoot, 81920)), ".data", limit, null, token).ConfigureAwait(false);
            Directory.CreateDirectory(Path.GetDirectoryName(target)!);
            using (var input = new FileStream(captured.FilePath, FileMode.Open, FileAccess.Read, FileShare.Read))
            using (var output = OfficeTemporaryFile.CreateAtPath(target, 81920, FileOptions.Asynchronous))
                await input.CopyToAsync(output, 81920, token).ConfigureAwait(false);
            bytes = checked(bytes + captured.Length);
            Guard.Files.Add(source, captured.Fingerprint);
        }

        private static bool IsRegularEntry(string path, bool directory) {
            if (!File.Exists(path) && !Directory.Exists(path)) return false;
            FileAttributes attributes = File.GetAttributes(path);
            return (attributes & FileAttributes.ReparsePoint) == 0 && ((attributes & FileAttributes.Directory) != 0) == directory;
        }

        private static string? FindToc(string directory) {
            if (!IsRegularEntry(directory, directory: true)) return null;
            string conventional = Path.Combine(directory, "Open Notebook.onetoc2");
            if (IsRegularEntry(conventional, directory: false)) return conventional;
            return Directory.EnumerateFiles(directory, "*.onetoc2").Where(path => IsRegularEntry(path, directory: false))
                .OrderBy(path => path, StringComparer.OrdinalIgnoreCase).FirstOrDefault();
        }

        private static Task<string> FingerprintAsync(string path, long maximumBytes, CancellationToken token) =>
            OfficeStreamPublication.ReadFingerprintAsync(_ => Task.FromResult<Stream>(
                OfficePathIdentity.OpenRegularFileForRead(path, Path.GetDirectoryName(path)!, 81920)), maximumBytes, token);

        internal OneNoteNotebook Read() {
            OneNoteNotebook notebook = OneNoteNotebookReader.Read(TableOfContentsPath, options);
            string root = Path.GetDirectoryName(TableOfContentsPath)!;
            string? Original(string? path) => path != null && (path == root || path.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.Ordinal))
                ? Guard.SourceRoot + path.Substring(root.Length) : path;
            void Restore(OneNoteSection section) {
                section.SourcePath = Original(section.SourcePath);
                foreach (OneNoteDiagnostic diagnostic in section.Diagnostics) diagnostic.SourcePath = Original(diagnostic.SourcePath);
            }
            notebook.SourcePath = Guard.SourceRoot;
            notebook.Name = new DirectoryInfo(Guard.SourceRoot).Name;
            foreach (OneNoteDiagnostic diagnostic in notebook.Diagnostics) diagnostic.SourcePath = Original(diagnostic.SourcePath);
            foreach (OneNoteSection section in notebook.Sections) Restore(section);
            var groups = new Stack<OneNoteSectionGroup>(notebook.SectionGroups);
            while (groups.Count > 0) {
                OneNoteSectionGroup group = groups.Pop();
                foreach (OneNoteSection section in group.Sections) Restore(section);
                foreach (OneNoteSectionGroup child in group.SectionGroups) groups.Push(child);
            }
            return notebook;
        }

        public void Dispose() => Directory.Delete(temporaryRoot, recursive: true);

        internal sealed class OneNoteNotebookGuard(string sourceRoot, long maximumBytes) : IOfficePublicationGuard {
            internal string SourceRoot { get; } = sourceRoot;
            internal Dictionary<string, string?> Files { get; } = new(StringComparer.Ordinal);
            internal Dictionary<string, (bool Exists, string? Toc)> Groups { get; } = new(StringComparer.Ordinal);

            internal async Task VerifyAsync(CancellationToken token) {
                foreach (var group in Groups) {
                    token.ThrowIfCancellationRequested();
                    if (group.Value.Exists != IsRegularEntry(group.Key, directory: true) ||
                        !string.Equals(group.Value.Toc, FindToc(group.Key), StringComparison.Ordinal))
                        throw new IOException("The OneNote notebook hierarchy changed; no output was published.");
                }
                foreach (var file in Files) {
                    token.ThrowIfCancellationRequested();
                    if (file.Value == null) {
                        if (IsRegularEntry(file.Key, directory: false)) throw new IOException("A OneNote section appeared during conversion.");
                    } else await OfficeStreamPublication.VerifyFingerprintAsync(_ => Task.FromResult<Stream>(
                        OfficePathIdentity.OpenRegularFileForRead(file.Key, SourceRoot, 81920)), file.Value, maximumBytes, token).ConfigureAwait(false);
                }
            }

            public async Task<bool> CanPublishAsync(string path, bool isDirectory, CancellationToken cancellationToken) {
                await VerifyAsync(cancellationToken).ConfigureAwait(false);
                return !Files.Keys.Any(source => OfficePathIdentity.AreEquivalent(path, source) ||
                    isDirectory && OfficePathIdentity.IsSameOrDescendant(source, path));
            }
        }
    }
}
