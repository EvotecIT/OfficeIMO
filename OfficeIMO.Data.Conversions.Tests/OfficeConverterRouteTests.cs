using System.Text.Json;
using OfficeIMO.Data.Conversions;
using Xunit;

namespace OfficeIMO.Data.Conversions.Tests;

// The front door runs the catalog's data routes by file extension, with FailOnLoss and data options carried through.
public class OfficeConverterRouteTests {
    [Fact]
    public void FileRoutesConvertByExtension() {
        string directory = Path.Combine(Path.GetTempPath(), "officeimo-data-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string yaml = Path.Combine(directory, "in.yaml");
            File.WriteAllText(yaml, "items:\n  - id: 1\n    name: a\n  - id: 2\n    name: b\n");
            OfficeConverter.Convert(yaml, Path.Combine(directory, "out.json"));
            using (JsonDocument json = JsonDocument.Parse(File.ReadAllBytes(Path.Combine(directory, "out.json")))) {
                Assert.Equal(2, json.RootElement.GetProperty("items").GetArrayLength());
            }

            OfficeConverter.Convert(Path.Combine(directory, "out.json"), Path.Combine(directory, "back.yaml"));
            Assert.StartsWith("items:", File.ReadAllText(Path.Combine(directory, "back.yaml")));
            OfficeConverter.Convert(Path.Combine(directory, "out.json"), Path.Combine(directory, "out.xlsx"));
            OfficeConverter.Convert(Path.Combine(directory, "out.xlsx"), Path.Combine(directory, "back.jsonl"));
            Assert.Equal(2, File.ReadAllLines(Path.Combine(directory, "back.jsonl")).Length);

            File.WriteAllText(Path.Combine(directory, "note.yaml"), "# comment\na: 1\n");
            Assert.Throws<OfficeConversionException>(() => OfficeConverter.Convert(Path.Combine(directory, "note.yaml"), Path.Combine(directory, "note.json"),
                new OfficeConversionOptions { FailOnLoss = true }));
            Assert.False(File.Exists(Path.Combine(directory, "note.json")));
        } finally {
            Directory.Delete(directory, recursive: true);
        }
    }
}
