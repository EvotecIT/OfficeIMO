using OfficeIMO.Core.Internal;
using System.Text.Json;

// Opt-in byte evidence for producer-controlled differential fixtures.
internal static class NativeLayoutDump {
    internal static int Run(string input, string output) {
        if (Directory.Exists(output)) throw new IOException("Choose a new output directory.");
        var bytes = File.ReadAllBytes(input);
        var limits = new OfficeCompoundReadOptions(4096, 2048, 32 * 1024 * 1024, 64 * 1024 * 1024);
        if (!OfficeCompoundFileReader.TryRead(bytes, limits, out var file, out var error) || file == null)
            throw new InvalidDataException(error);
        Directory.CreateDirectory(output);
        foreach (var stream in file.Streams) {
            // Retain names as readable filenames without platform-sensitive separators.
            var name = string.Concat(stream.Key.Select(c => char.IsControl(c) || c == '/' || c == '\\' ? '_' : c));
            File.WriteAllBytes(Path.Combine(output, name), stream.Value);
        }
        Console.WriteLine(JsonSerializer.Serialize(file.Streams.Select(s => new { name = s.Key, bytes = s.Value.Length })));
        return 0;
    }
}
