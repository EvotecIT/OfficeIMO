using Xunit;

namespace OfficeIMO.Yaml.Tests.Serialization;

// Contract: generated and reflection metadata are two sources for one serializer, so every model writes the same YAML
// through both and reads back to a value that writes the same YAML again (no member, type or value is lost on the way).
public class YamlSerializerRoundTripTests {
    private const int Iterations = 150;
    private static readonly YamlSerializerOptions Reflection = new YamlSerializerOptions();

    [Fact]
    public void GeneratedAndReflectionMetadataWriteTheSameYamlAndReadItBack() {
        TestYamlContext context = TestYamlContext.Default;
        for (int seed = 0; seed < Iterations; seed++) {
            var random = new Random(seed);
            Check(SerializerSamples.Config(random), context.AppConfig, seed);
            Check(SerializerSamples.Scalars(random), context.Scalars, seed);
            Check(SerializerSamples.Collections(random), context.Collections, seed);
            Check(SerializerSamples.Zoo(random), context.Zoo, seed);
            Check(SerializerSamples.Kube(random), context.KubeObject, seed);
            Check(SerializerSamples.People(random), context.ListPerson, seed);
            Check(SerializerSamples.Account(random), context.Account, seed);
        }
    }

    [Fact]
    public void RecordsReadBackEqual() {
        var random = new Random(7);
        for (int i = 0; i < Iterations; i++) {
            List<Person> people = SerializerSamples.People(random);
            string yaml = YamlSerializer.Serialize(people, TestYamlContext.Default.ListPerson);
            Assert.Equal(people, YamlSerializer.Deserialize(yaml, TestYamlContext.Default.ListPerson));
            Assert.Equal(people, YamlSerializer.Deserialize<List<Person>>(yaml, Reflection));
        }
    }

    private static void Check<T>(T value, YamlTypeInfo<T> generated, int seed) {
        string viaGenerated = YamlSerializer.Serialize(value, generated);
        string viaReflection = YamlSerializer.Serialize(value, Reflection);
        Assert.True(viaGenerated == viaReflection, "Seed " + seed + ": generated and reflection output differ.\n" + viaGenerated + "\n----\n" + viaReflection);

        T? fromGenerated = YamlSerializer.Deserialize(viaGenerated, generated);
        T? fromReflection = YamlSerializer.Deserialize<T>(viaGenerated, Reflection);
        string again = YamlSerializer.Serialize(fromGenerated, generated!);
        Assert.True(viaGenerated == again, "Seed " + seed + ": generated metadata lost data.\n" + viaGenerated + "\n----\n" + again);
        Assert.Equal(viaGenerated, YamlSerializer.Serialize(fromReflection, Reflection));

        // Independent parse: the text is valid YAML that the document model loads.
        YamlDocument.Parse(viaGenerated);
    }
}
