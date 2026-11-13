using OfficeIMO.Yaml.Tests.Conformance;
using Xunit;
using DotNet = YamlDotNet;
using Sharp = SharpYaml;

namespace OfficeIMO.Yaml.Tests.Serialization;

// Contract: what the serializer writes is ordinary YAML that independent libraries read into the same objects, and real
// files read into the same objects through all three where their type rules agree. Objects are compared through this
// serializer's deterministic output, which the round-trip tests prove lossless.
public class YamlSerializerInteropTests {
    private static readonly TestYamlContext Pascal = TestYamlContext.Default;
    private static readonly TestYamlContext Camel = new TestYamlContext(new YamlSerializerOptions { PropertyNamingPolicy = OfficeNamingPolicy.CamelCase });

    private static readonly Dictionary<string, Type> Animals = new Dictionary<string, Type> { ["!cat"] = typeof(Cat), ["!dog"] = typeof(Dog) };
    private static readonly Dictionary<string, Type> Kinds = new Dictionary<string, Type> { ["Deployment"] = typeof(KubeDeployment), ["Service"] = typeof(KubeService) };

    [Fact]
    public void OurOutputReadsIntoEqualObjectsWithYamlDotNetAndSharpYaml() {
        DotNet.Serialization.IDeserializer dotNet = new DotNet.Serialization.DeserializerBuilder()
            .WithTagMapping("!cat", typeof(Cat))
            .WithTagMapping("!dog", typeof(Dog))
            .Build();
        var sharp = new Sharp.YamlSerializerOptions {
            PolymorphismOptions = new Sharp.YamlPolymorphismOptions {
                DiscriminatorStyle = Sharp.YamlTypeDiscriminatorStyle.Tag,
                DerivedTypeMappings = {
                    [typeof(Animal)] = new List<Sharp.YamlDerivedType> { new Sharp.YamlDerivedType(typeof(Cat), "cat") { Tag = "!cat" }, new Sharp.YamlDerivedType(typeof(Dog), "dog") { Tag = "!dog" } }
                }
            }
        };

        DotNet.Serialization.IDeserializer kubeDotNet = KubeDotNet();
        Sharp.YamlSerializerOptions kubeSharp = KubeSharp();
        for (int seed = 0; seed < 60; seed++) {
            var random = new Random(seed);
            KubeObject kube = SerializerSamples.Kube(random);
            string manifest = YamlSerializer.Serialize(kube, Camel.KubeObject);
            Assert.Equal(manifest, YamlSerializer.Serialize(kubeDotNet.Deserialize<KubeObject>(manifest), Camel.KubeObject));
            AssertSharpYamlReadsAlike(manifest, Camel.KubeObject, kubeSharp);

            AppConfig config = SerializerSamples.Config(random);
            // The libraries' rules for Uri differ: YamlDotNet 18 reads a null scalar as the text "null", and SharpYaml
            // 3.17 has no Uri scalar, so each reads the member only where it agrees.
            config.Endpoint = new Uri("relative/path", UriKind.Relative);
            string yaml = YamlSerializer.Serialize(config, Pascal.AppConfig);
            Assert.Equal(yaml, YamlSerializer.Serialize(dotNet.Deserialize<AppConfig>(yaml), Pascal.AppConfig));
            config.Endpoint = null;
            AssertSharpYamlReadsAlike(YamlSerializer.Serialize(config, Pascal.AppConfig), Pascal.AppConfig, sharp);

            Zoo zoo = SerializerSamples.Zoo(random);
            yaml = YamlSerializer.Serialize(zoo, Pascal.Zoo);
            Assert.Equal(yaml, YamlSerializer.Serialize(dotNet.Deserialize<Zoo>(yaml), Pascal.Zoo));
            AssertSharpYamlReadsAlike(yaml, Pascal.Zoo, sharp);
        }
    }

    [Fact]
    public void KubernetesStreamReadsAlikeInAllThreeLibraries() {
        string text = File.ReadAllText(Path.Combine(YamlAssetPaths.Root, "Corpus", "Kubernetes", "guestbook-all-in-one.yaml"));

        var ours = new List<KubeObject>();
        using (var reader = new YamlReader(text)) {
            while (reader.TokenType != YamlTokenType.StreamEnd) {
                KubeObject? document = YamlSerializer.Deserialize(reader, Camel.KubeObject);
                if (document is not null) ours.Add(document);
            }
        }
        Assert.Equal(6, ours.Count);
        var deployment = (KubeDeployment)ours[1];
        Assert.Equal(("redis-leader", 1, "100Mi"), (deployment.Metadata.Name, deployment.Spec.Replicas, deployment.Spec.Template.Spec.Containers[0].Resources!.Requests!["memory"]));

        DotNet.Serialization.IDeserializer dotNet = KubeDotNet();
        var parser = new DotNet.Core.Parser(new StringReader(text));
        DotNet.Core.ParserExtensions.Consume<DotNet.Core.Events.StreamStart>(parser);
        var theirs = new List<KubeObject>();
        while (DotNet.Core.ParserExtensions.Accept<DotNet.Core.Events.DocumentStart>(parser, out _)) theirs.Add(dotNet.Deserialize<KubeObject>(parser));
        AssertSame(ours, theirs);

        Sharp.YamlSerializerOptions sharp = KubeSharp();
        var sharpDocuments = new List<KubeObject>();
        foreach (string document in System.Text.RegularExpressions.Regex.Split(text, "^---[^\\n]*\\n", System.Text.RegularExpressions.RegexOptions.Multiline)) {
            if (document.Trim().Length > 0) sharpDocuments.Add(Sharp.YamlSerializer.Deserialize<KubeObject>(document, sharp)!);
        }
        AssertSame(ours, sharpDocuments);
    }

    [Fact]
    public void ComposeAnchorsAndMergeKeysReadAlikeWithYamlDotNet() {
        string text = File.ReadAllText(Path.Combine(YamlAssetPaths.Root, "Corpus", "Compose", "airflow-docker-compose.yaml"));
        ComposeFile ours = YamlSerializer.Deserialize(text, Camel.ComposeFile)!;
        ComposeService worker = ours.Services["airflow-worker"];
        Assert.Equal("celery worker", worker.Command);
        Assert.Equal("CeleryExecutor", worker.Environment!["AIRFLOW__CORE__EXECUTOR"]);
        Assert.Equal("0", worker.Environment["DUMB_INIT_SETSID"]);
        Assert.Equal(4, worker.Volumes!.Count);

        DotNet.Serialization.IDeserializer dotNet = new DotNet.Serialization.DeserializerBuilder()
            .WithNamingConvention(DotNet.Serialization.NamingConventions.CamelCaseNamingConvention.Instance)
            .IgnoreUnmatchedProperties()
            .Build();
        ComposeFile theirs = dotNet.Deserialize<ComposeFile>(new DotNet.Core.MergingParser(new DotNet.Core.Parser(new StringReader(text))));
        Assert.Equal(YamlSerializer.Serialize(ours, Camel.ComposeFile), YamlSerializer.Serialize(Normalize(theirs), Camel.ComposeFile));

        // SharpYaml 3.17 rejects this file's '<<: *airflow-common' ("Merge alias values are not supported unless they
        // resolve to a mapping"), although the anchor names a mapping, so its rules do not agree here.
        var sharp = new Sharp.YamlSerializerOptions { PropertyNamingPolicy = System.Text.Json.JsonNamingPolicy.CamelCase };
        Assert.Throws<Sharp.YamlException>(() => Sharp.YamlSerializer.Deserialize<ComposeFile>(text, sharp));
    }

    /// <summary>
    /// Compares SharpYaml's reading where its rules agree. SharpYaml 3.17 reads a quoted null spelling ('Null', '~', '')
    /// inside a polymorphic type as null, and a quoted '&lt;&lt;' key as a merge key; YAML makes both strings, so those
    /// scalars are renamed before the comparison.
    /// </summary>
    private static void AssertSharpYamlReadsAlike<T>(string yaml, YamlTypeInfo<T> typeInfo, Sharp.YamlSerializerOptions options) {
        int renamed = 0;
        string comparable = System.Text.RegularExpressions.Regex.Replace(yaml, "(?<=(?:^|[:-] ) *)'(null|Null|NULL|~|<<)?'(?=:|\\n)",
            _ => "'renamed-" + (++renamed).ToString(System.Globalization.CultureInfo.InvariantCulture) + "'", System.Text.RegularExpressions.RegexOptions.Multiline);
        string expected = YamlSerializer.Serialize(YamlSerializer.Deserialize(comparable, typeInfo)!, typeInfo);
        Assert.Equal(expected, YamlSerializer.Serialize(Sharp.YamlSerializer.Deserialize<T>(comparable, options)!, typeInfo));
    }

    private static DotNet.Serialization.IDeserializer KubeDotNet() => new DotNet.Serialization.DeserializerBuilder()
        .WithNamingConvention(DotNet.Serialization.NamingConventions.CamelCaseNamingConvention.Instance)
        .IgnoreUnmatchedProperties()
        .WithTypeDiscriminatingNodeDeserializer(options => options.AddKeyValueTypeDiscriminator<KubeObject>("kind", Kinds))
        .Build();

    private static Sharp.YamlSerializerOptions KubeSharp() => new Sharp.YamlSerializerOptions {
        PropertyNamingPolicy = System.Text.Json.JsonNamingPolicy.CamelCase,
        PolymorphismOptions = new Sharp.YamlPolymorphismOptions {
            DiscriminatorStyle = Sharp.YamlTypeDiscriminatorStyle.Property,
            TypeDiscriminatorPropertyName = "kind",
            DerivedTypeMappings = {
                [typeof(KubeObject)] = new List<Sharp.YamlDerivedType> { new Sharp.YamlDerivedType(typeof(KubeDeployment), "Deployment"), new Sharp.YamlDerivedType(typeof(KubeService), "Service") }
            }
        }
    };

    private static void AssertSame(List<KubeObject> expected, List<KubeObject> actual) {
        Assert.Equal(expected.Count, actual.Count);
        for (int i = 0; i < expected.Count; i++) {
            Assert.Equal(YamlSerializer.Serialize(expected[i], Camel.KubeObject), YamlSerializer.Serialize(actual[i], Camel.KubeObject));
        }
    }

    /// <summary>Other libraries read untyped sequences as their own list types; this serializer writes any list alike.</summary>
    private static ComposeFile Normalize(ComposeFile file) {
        foreach (ComposeService service in file.Services.Values) {
            if (service.Command is System.Collections.IEnumerable items and not string) service.Command = items.Cast<object?>().ToList();
        }
        return file;
    }
}
