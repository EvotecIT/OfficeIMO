using System.Numerics;

namespace OfficeIMO.Yaml.Tests.Serialization;

/// <summary>Seeded generators for the representative models, mixing ordinary values with the ones YAML gets wrong.</summary>
internal static class SerializerSamples {
    /// <summary>Strings a YAML 1.1 or 1.2 reader could take for another type, or that need quoting or escapes.</summary>
    public static readonly string[] TrickyStrings = {
        "", " ", "yes", "No", "on", "OFF", "y", "n", "~", "null", "Null", "true", "False", "1e3", "0o17", "012", "0x1F", "1:20",
        ".inf", "-.NaN", "2001-12-14", "<<", "=", " padded ", "a: b", "#tag", "- item", "- ", "? key", "1_000", "+12", "1.0", "-0",
        "1.29.2", "10.0.0.1", "line one\nline two", "trailing newline\n", "tab\there", "quote ' and \" both", "back\\slash",
        "unicode: zażółć gęślą jaźń", "emoji 🎉", "control \u0001 char", "&anchor", "*alias", "!tag", "|", ">", "%directive",
        "@at", "`tick", "{flow}", "[seq]", "comma, separated", "---", "...", "key:", "multi\n\nblank\n", "\u2028separator"
    };

    public static string Text(Random random) => random.Next(3) == 0
        ? TrickyStrings[random.Next(TrickyStrings.Length)]
        : "value-" + random.Next(1000).ToString(System.Globalization.CultureInfo.InvariantCulture);

    public static string? NullableText(Random random) => random.Next(4) == 0 ? null : Text(random);

    public static AppConfig Config(Random random) {
        var config = new AppConfig {
            Name = Text(random),
            Port = random.Next(-1, 65536),
            Enabled = random.Next(2) == 0,
            Level = (LogLevel)random.Next(5),
            Access = (Permissions)random.Next(8),
            Ratio = random.Next(6) switch { 0 => double.NaN, 1 => double.PositiveInfinity, 2 => -0.0, 3 => 1e-300, _ => random.NextDouble() * 1000 },
            Price = (decimal)random.Next(100000) / 100m,
            Timeout = TimeSpan.FromMilliseconds(random.Next(0, 100_000_000)),
            Endpoint = random.Next(3) == 0 ? null : new Uri(random.Next(2) == 0 ? "https://example.com/api?q=1" : "relative/path", UriKind.RelativeOrAbsolute),
            Retries = random.Next(3) == 0 ? null : random.Next(10),
            Comment = NullableText(random),
            Hosts = Enumerable.Range(0, random.Next(4)).Select(_ => Text(random)).ToArray()
        };
        for (int i = random.Next(5); i > 0; i--) config.Tags.Add(Text(random));
        for (int i = random.Next(4); i > 0; i--) config.Labels["k" + i + Text(random)] = Text(random);
        for (int i = random.Next(3); i > 0; i--) config.Features[Text(random)] = random.Next(2) == 0;
        return config;
    }

    public static Scalars Scalars(Random random) {
        var bytes = new byte[random.Next(0, 40)];
        random.NextBytes(bytes);
        return new Scalars {
            Text = Text(random),
            Letter = "a:#'\" \n€"[random.Next(8)],
            Byte = (byte)random.Next(256),
            SByte = (sbyte)random.Next(-128, 128),
            Short = (short)random.Next(short.MinValue, short.MaxValue + 1),
            UShort = (ushort)random.Next(ushort.MaxValue + 1),
            UInt = random.Next(2) == 0 ? uint.MaxValue : (uint)random.Next(),
            Long = random.Next(3) switch { 0 => long.MinValue, 1 => long.MaxValue, _ => random.Next() * (long)random.Next(-5, 5) },
            ULong = random.Next(2) == 0 ? ulong.MaxValue : (ulong)random.Next(),
            Single = random.Next(4) switch { 0 => 0.1f, 1 => float.MaxValue, 2 => float.NegativeInfinity, _ => (float)random.NextDouble() },
            Double = random.Next(4) switch { 0 => 0.1, 1 => double.MinValue, 2 => double.Epsilon, _ => random.NextDouble() * 1e10 },
            Decimal = random.Next(3) switch { 0 => decimal.MaxValue, 1 => 1.50m, _ => -(decimal)random.NextDouble() },
            Big = BigInteger.Pow(new BigInteger(random.Next(2, 100)), random.Next(1, 40)) * (random.Next(2) == 0 ? -1 : 1),
            When = new DateTime(2000 + random.Next(30), 1 + random.Next(12), 1 + random.Next(28), random.Next(24), random.Next(60), random.Next(60),
                (DateTimeKind)random.Next(2) == DateTimeKind.Unspecified ? DateTimeKind.Unspecified : DateTimeKind.Utc).AddTicks(random.Next(2) == 0 ? 0 : random.Next(10_000_000)),
            At = new DateTimeOffset(2020, 2, 29, 23, 59, 59, TimeSpan.FromMinutes(random.Next(-12, 13) * 60)).AddTicks(random.Next(10_000)),
            Duration = random.Next(2) == 0 ? TimeSpan.FromHours(1.5) : -TimeSpan.FromTicks(random.Next()),
            Id = new Guid(Enumerable.Range(0, 16).Select(_ => (byte)random.Next(256)).ToArray()),
            Payload = bytes,
            Origin = new Point { X = random.Next(-100, 100), Y = random.Next(-100, 100) },
            Maybe = random.Next(2) == 0 ? null : new Point { X = 1, Y = -1 },
            Box = new Size(random.Next(1000), random.Next(1000)),
#if NET6_0_OR_GREATER
            Day = new DateOnly(1999 + random.Next(30), 1 + random.Next(12), 1 + random.Next(28)),
            Clock = new TimeOnly(random.Next(24), random.Next(60), random.Next(60), random.Next(1000)),
#endif
        };
    }

    public static Collections Collections(Random random) {
        var collections = new Collections {
            Numbers = Enumerable.Range(0, random.Next(5)).Select(_ => random.Next(-1000, 1000)).ToArray(),
            ReadOnly = Enumerable.Range(0, random.Next(4)).Select(_ => Text(random)).ToList(),
            Sequence = Enumerable.Range(0, random.Next(4)).Select(_ => (long)random.Next()).ToList(),
            Unique = new HashSet<string>(Enumerable.Range(0, random.Next(4)).Select(i => "u" + i)),
            Set = new HashSet<int>(Enumerable.Range(0, random.Next(4))),
            Linked = new LinkedList<string>(Enumerable.Range(0, random.Next(3)).Select(_ => Text(random))),
            Anything = random.Next(3) switch {
                0 => null,
                1 => new Dictionary<string, object?> { ["a"] = 1L, ["b"] = new List<object?> { "x", true, 2.5, null } },
                _ => "plain"
            }
        };
        for (int i = random.Next(4); i > 0; i--) collections.Sorted["s" + i] = random.Next();
        var byNumber = new Dictionary<int, string>();
        for (int i = random.Next(4); i > 0; i--) byNumber[random.Next(-50, 50)] = Text(random);
        collections.ByNumber = byNumber;
        for (int i = random.Next(3); i > 0; i--) collections.ByLevel[(LogLevel)random.Next(5)] = new List<string> { Text(random) };
        for (int i = random.Next(3); i > 0; i--) collections.Nested.Add(new Dictionary<string, Point> { ["p" + i] = new Point { X = i, Y = -i } });
        return collections;
    }

    public static Zoo Zoo(Random random) {
        var zoo = new Zoo();
        for (int i = random.Next(5); i > 0; i--) {
            zoo.Animals.Add(random.Next(2) == 0 ? new Cat { Name = Text(random), Lives = random.Next(10) } : new Dog { Name = Text(random), GoodBoy = random.Next(2) == 0 });
        }
        zoo.Favorite = zoo.Animals.Count > 0 && random.Next(2) == 0 ? new Cat { Name = "Tom", Lives = 9 } : null;
        return zoo;
    }

    public static KubeObject Kube(Random random) {
        var metadata = new KubeMetadata {
            Name = "app-" + random.Next(100),
            Namespace = random.Next(2) == 0 ? null : "prod",
            Labels = random.Next(3) == 0 ? null : new Dictionary<string, string> { ["app"] = "web", ["tier"] = Text(random) },
            Annotations = random.Next(3) == 0 ? new Dictionary<string, string> { ["note"] = Text(random) } : null
        };
        if (random.Next(2) == 0) {
            return new KubeService {
                ApiVersion = "v1",
                Metadata = metadata,
                Spec = new ServiceSpec {
                    Type = random.Next(2) == 0 ? null : "ClusterIP",
                    Ports = { new ServicePort { Port = 80, TargetPort = random.Next(2) == 0 ? null : 8080, Protocol = "TCP" } },
                    Selector = new Dictionary<string, string> { ["app"] = "web" }
                }
            };
        }
        var deployment = new KubeDeployment {
            ApiVersion = "apps/v1",
            Metadata = metadata,
            Spec = new DeploymentSpec {
                Replicas = random.Next(3) == 0 ? null : random.Next(10),
                Selector = new LabelSelector { MatchLabels = { ["app"] = "web" } }
            }
        };
        for (int i = random.Next(1, 4); i > 0; i--) {
            deployment.Spec.Template.Spec.Containers.Add(new Container {
                Name = "c" + i,
                Image = "registry.example.com/app:1." + random.Next(30) + ".0",
                ImagePullPolicy = random.Next(2) == 0 ? null : "IfNotPresent",
                Args = random.Next(2) == 0 ? null : new List<string> { "--port", "8080", Text(random) },
                Ports = new List<PortSpec> { new PortSpec { ContainerPort = 8080, Name = "http" } },
                Env = new List<EnvVar> { new EnvVar { Name = "MODE", Value = Text(random) }, new EnvVar { Name = "EMPTY" } },
                Resources = new ResourceRequirements { Requests = new Dictionary<string, string> { ["cpu"] = "100m", ["memory"] = "128Mi" } }
            });
        }
        return deployment;
    }

    public static List<Person> People(Random random) =>
        Enumerable.Range(0, random.Next(6)).Select(_ => new Person(Text(random), random.Next(120), random.Next(2) == 0 ? null : Text(random))).ToList();

    public static Account Account(Random random) => new Account {
        Id = "acc-" + random.Next(1000),
        Owner = Text(random),
        Balance = random.Next(-1000, 1000) / 10m,
        Opened = new DateTimeOffset(2024, 1 + random.Next(12), 1, 9, 30, 0, TimeSpan.FromHours(2))
    };
}
