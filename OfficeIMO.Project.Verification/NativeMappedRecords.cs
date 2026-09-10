using OfficeIMO.Core.Internal;
using OfficeIMO.Project;
using System.Text.Json;

internal static class NativeMappedRecords {
    internal static int Run(string input, string output) {
        if (Directory.Exists(output)) throw new IOException("Choose a new output directory.");
        var bytes = File.ReadAllBytes(input);
        if (!OfficeCompoundFileReader.TryRead(bytes, new OfficeCompoundReadOptions(4096, 2048, 32 * 1024 * 1024, 64 * 1024 * 1024), out var file, out var error) || file == null)
            throw new InvalidDataException(error);
        var properties = ProjectNativeProperties.Read(file.Streams["   114/Props"], default);
        var result = new Dictionary<string, object>();
        foreach (var spec in new[] { ("Task", 0x14u, 0x0b400056u), ("Rsc", 0x15u, 0x0c40001bu), ("Cal", 0x16u, 0x0d400009u), ("Assn", 0x17u, 0x0f400000u), ("Cons", 0x18u, 0x0e400000u) }) {
            var table = new ProjectNativeTable(file, "TBknd" + spec.Item1, properties[0x03000000u | spec.Item2], properties[0x00020000u | spec.Item2], 300000, default);
            var records = new List<object>();
            foreach (var record in table.Records) {
                record.Uid = record.Integer(spec.Item3) ?? throw new InvalidDataException("Missing UID field: " + spec.Item1);
                var values = new Dictionary<string, object>();
                foreach (var field in table.Fields.Values) {
                    if (record.Boolean(field.Id) is bool flag) {
                        values.Add(field.Id.ToString("X8"), new { field.Type, field.Source, field.Secondary, bytes = 1, hex = flag ? "01" : "00" });
                        continue;
                    }
                    var value = record.Value(field.Id);
                    if (!value.HasValue) continue;
                    values.Add(field.Id.ToString("X8"), new { field.Type, field.Source, field.Secondary, bytes = value.Value.Length,
                        hex = Convert.ToHexString(value.Value.Slice(0, Math.Min(value.Value.Length, 128)).Copy()) });
                }
                records.Add(new { record.Uid, values });
            }
            result.Add(spec.Item1, records);
        }
        Directory.CreateDirectory(output);
        File.WriteAllText(Path.Combine(output, "records.json"), JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true }));
        Console.WriteLine(JsonSerializer.Serialize(result.Keys));
        return 0;
    }
}
