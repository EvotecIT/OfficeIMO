using System.Reflection;

namespace OfficeIMO.Email.Store.Tests;

internal static class PstAllocationMapTestSupport {
    private const BindingFlags PrivateInstance = BindingFlags.Instance | BindingFlags.NonPublic;

    internal static T Field<T>(object owner, string name) =>
        (T)owner.GetType().GetField(name, PrivateInstance)!.GetValue(owner)!;

    internal static PstWriterFile WriterFile(EmailStorePstWriter writer) =>
        Field<PstWriterFile>(Field<PstStoreWriterCore>(writer, "_core"), "_file");

    internal static void RegisterThrough(PstWriterFile file, long end) =>
        typeof(PstWriterFile).GetMethod("RegisterMapPagesThrough", PrivateInstance)!.Invoke(file, new object[] { end });

    internal static CountingFileStream CountMapWrites(PstWriterFile file) {
        PstWriterAllocationMap map = Field<PstWriterAllocationMap>(file, "_allocationMap");
        FileStream original = Field<FileStream>(map, "_stream");
        string path = original.Name;
        original.Dispose();
        var counted = new CountingFileStream(path);
        // Count the real map-file writes without adding an instrumentation API to the library.
        typeof(PstWriterAllocationMap).GetField("_stream", PrivateInstance)!.SetValue(map, counted);
        return counted;
    }

    internal sealed class CountingFileStream : FileStream {
        internal long WriteCalls { get; private set; }

        internal CountingFileStream(string path) : base(path, FileMode.Open, FileAccess.ReadWrite,
            FileShare.Read, 16 * 1024, FileOptions.RandomAccess) { }

        public override void Write(byte[] buffer, int offset, int count) {
            WriteCalls++;
            base.Write(buffer, offset, count);
        }
    }
}
