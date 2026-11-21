using Xunit;

namespace OfficeIMO.Toml.Tests.Serialization;

// Contract: generated and reflection metadata are two sources for one serializer, so every model writes the same TOML
// through both and reads back to a value that writes the same TOML again (no member, type or value is lost on the way).
// Tomlyn, an independent TOML 1.1 parser, must read every output to the same values the document model sees.
public class TomlSerializerRoundTripTests {
    private const int Iterations = 120;
    private static readonly TomlSerializerOptions Reflection = new TomlSerializerOptions();

    [Fact]
    public void GeneratedAndReflectionMetadataWriteTheSameTomlAndReadItBack() {
        TestTomlContext context = TestTomlContext.Default;
        for (int seed = 0; seed < Iterations; seed++) {
            var random = new Random(seed);
            Check(TomlSerializerSamples.Cargo(random), context.CargoManifest, seed);
            Check(TomlSerializerSamples.Scalars(random), context.Scalars, seed);
            Check(TomlSerializerSamples.Collections(random), context.Collections, seed);
            Check(TomlSerializerSamples.Team(random), context.Team, seed);
            Check(TomlSerializerSamples.Monitoring(random), context.Monitoring, seed);
        }
    }

    [Fact]
    public void RecordsReadBackEqual() {
        var random = new Random(7);
        for (int i = 0; i < Iterations; i++) {
            Team team = TomlSerializerSamples.Team(random);
            string toml = TomlSerializer.Serialize(team, TestTomlContext.Default.Team);
            Assert.Equal(team.Members, TomlSerializer.Deserialize(toml, TestTomlContext.Default.Team).Members);
            Assert.Equal(team.Members, TomlSerializer.Deserialize<Team>(toml, Reflection).Members);
        }
    }

    private static void Check<T>(T value, TomlTypeInfo<T> generated, int seed) {
        string viaGenerated = TomlSerializer.Serialize(value, generated);
        string viaReflection = TomlSerializer.Serialize(value, Reflection);
        Assert.True(viaGenerated == viaReflection, "Seed " + seed + ": generated and reflection output differ.\n" + viaGenerated + "\n----\n" + viaReflection);

        T fromGenerated = TomlSerializer.Deserialize(viaGenerated, generated);
        T fromReflection = TomlSerializer.Deserialize<T>(viaGenerated, Reflection);
        string again = TomlSerializer.Serialize(fromGenerated, generated);
        Assert.True(viaGenerated == again, "Seed " + seed + ": generated metadata lost data.\n" + viaGenerated + "\n----\n" + again);
        Assert.Equal(viaGenerated, TomlSerializer.Serialize(fromReflection, Reflection));

        // Independent parse: Tomlyn reads the text to the same tree of values as the document model.
        var independent = global::Tomlyn.TomlSerializer.Deserialize<global::Tomlyn.Model.TomlTable>(viaGenerated)!;
        Assert.True(TomlInterop.Same(TomlDocument.Parse(viaGenerated).Root, independent, out string difference),
            "Seed " + seed + ": Tomlyn reads " + difference + "\n" + viaGenerated);
    }
}
