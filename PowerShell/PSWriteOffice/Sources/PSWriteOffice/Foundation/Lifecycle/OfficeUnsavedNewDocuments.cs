using System.Management.Automation;
using System.Runtime.CompilerServices;

namespace PSWriteOffice.Foundation;

/// <summary>Remembers where a live New- document will be saved, so closing it without a save can warn.</summary>
internal static class OfficeUnsavedNewDocuments {
    private sealed class Destination {
        internal Destination(string path) {
            Path = path;
            OriginalWriteTime = File.Exists(path) ? File.GetLastWriteTimeUtc(path) : null;
        }
        internal string Path { get; }
        internal DateTime? OriginalWriteTime { get; }
    }

    private static readonly ConditionalWeakTable<object, Destination> Destinations = new();

    internal static void Track(object document, string path) {
        Destinations.Remove(document);
        Destinations.Add(document, new Destination(path));
    }

    /// <summary>Warns when the destination given to New- was never written. Any save route counts, because the file itself is checked.</summary>
    internal static void WarnIfUnsaved(Cmdlet cmdlet, object document, string kind, string format) {
        if (!Destinations.TryGetValue(document, out Destination? destination) || destination == null) return;
        Destinations.Remove(document);
        bool saved = File.Exists(destination.Path) && File.GetLastWriteTimeUtc(destination.Path) != destination.OriginalWriteTime;
        if (saved) return;
        cmdlet.WriteWarning($"The {kind} was closed without being saved, so '{destination.Path}' was not written. " +
            $"Run Save-Office{format} before Close-Office{format}, or use New-Office{format} -Path <file> {{ ... }}, which saves when the script block ends.");
    }
}
