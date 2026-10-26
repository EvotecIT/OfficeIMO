using System;
using System.IO;
using System.Threading;
using OfficeIMO;
using OfficeIMO.Core.Internal;

namespace PSWriteOffice.Foundation;

/// <summary>Runs a document block and publishes its native save only after successful completion and loss validation.</summary>
internal static class OfficeContentLifecycle {
    /// <remarks>
    /// A new-document factory receives a private staging path. An open-document factory receives the input path
    /// and must disable autosave. Save writes only to the supplied staging path and returns the native-save report.
    /// The existing destination remains unchanged when authoring, save, validation or disposal fails.
    /// </remarks>
    internal static FileInfo Run<T>(string path, bool isNew, bool refuseExisting, bool allowLoss,
        Func<string, T> createOrLoad, Action<T> content, Func<T, string, bool, OfficeConversionReport?> save,
        CancellationToken cancellationToken = default) where T : class, IDisposable {
        OfficePath.ValidateDestination(path, refuseExisting);
        string directory = Path.GetDirectoryName(path)!;
        if (!Directory.Exists(directory)) throw new DirectoryNotFoundException("The output directory does not exist: " + directory);
        string stage = Path.Combine(directory, "." + Path.GetFileNameWithoutExtension(path) + "." + Guid.NewGuid().ToString("N") + Path.GetExtension(path));
        try {
            cancellationToken.ThrowIfCancellationRequested();
            using (T document = createOrLoad(isNew ? stage : path)) {
                using (OfficeDocumentContext.Push(document)) content(document);
                cancellationToken.ThrowIfCancellationRequested();
                var report = save(document, stage, allowLoss);
                if (!allowLoss) report?.ThrowIfLoss();
            }
            cancellationToken.ThrowIfCancellationRequested();
            OfficePath.ValidateDestination(path, refuseExisting);
            OfficeFileCommit.CommitTemporaryFileAtomically(stage, path,
                refuseExisting ? OfficeFileCommit.ConflictPolicy.FailIfExists : OfficeFileCommit.ConflictPolicy.Replace);
            return new FileInfo(path);
        } finally {
            if (File.Exists(stage)) File.Delete(stage);
        }
    }
}
