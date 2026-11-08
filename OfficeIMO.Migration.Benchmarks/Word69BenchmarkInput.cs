using System.Reflection;

namespace OfficeIMO.Migration.Benchmarks;

/// <summary>Untimed primitives shared by synthetic Word 6/95 benchmark inputs.</summary>
internal static class Word69BenchmarkInput {
    internal static byte[] WrapCompound(byte[] main) {
        // Use the canonical container writer without exposing fixture-generation APIs in product code.
        Assembly core = Assembly.Load("OfficeIMO.Core");
        Type streamType = core.GetType("OfficeIMO.Core.Internal.OfficeCompoundStream", throwOnError: true)!;
        Type writerType = core.GetType("OfficeIMO.Core.Internal.OfficeCompoundFileWriter", throwOnError: true)!;
        object stream = Activator.CreateInstance(streamType, BindingFlags.Instance | BindingFlags.NonPublic,
            binder: null, args: new object[] { "WordDocument", main }, culture: null)!;
        Array streams = Array.CreateInstance(streamType, 1);
        streams.SetValue(stream, 0);
        MethodInfo write = writerType.GetMethods(BindingFlags.Static | BindingFlags.NonPublic)
            .Single(method => method.Name == "Write" && method.ReturnType == typeof(byte[]) && method.GetParameters().Length == 2);
        return (byte[])write.Invoke(null, new object?[] { streams, null })!;
    }

    internal static void WriteWord(byte[] bytes, int offset, ushort value) {
        bytes[offset] = (byte)value;
        bytes[offset + 1] = (byte)(value >> 8);
    }

    internal static void WriteInt(byte[] bytes, int offset, int value) {
        WriteWord(bytes, offset, (ushort)value);
        WriteWord(bytes, offset + 2, (ushort)(value >> 16));
    }
}
