using System.Numerics;

namespace OfficeIMO.Toml.Tests.Serialization;

/// <summary>Seeded model instances that reach every member, with strings that need quoting and edge numbers.</summary>
internal static class TomlSerializerSamples {
    private static readonly string[] Words = { "serde", "tokio", "a b", "dotted.key", "quote\"d", "tab\there", "naïve", "", "line\nbreak", "#hash", "x = 1" };

    public static string Word(Random random) => Words[random.Next(Words.Length)];

    public static CargoManifest Cargo(Random random) {
        var manifest = new CargoManifest {
            Package = new CargoPackage {
                Name = Word(random),
                Version = random.Next(10) + "." + random.Next(10) + ".0",
                Edition = (Edition)random.Next(4),
                Authors = Enumerable.Range(0, random.Next(3)).Select(_ => Word(random)).ToList(),
                Description = random.Next(2) == 0 ? null : Word(random),
                RustVersion = random.Next(2) == 0 ? null : 70 + random.Next(20),
                Publish = random.Next(2) == 0
            }
        };
        for (int i = random.Next(4); i > 0; i--) {
            manifest.Dependencies[Word(random) + i] = new Dependency {
                Version = random.Next(3) == 0 ? null : "1." + i,
                Path = random.Next(3) == 0 ? "../crate" + i : null,
                Features = random.Next(2) == 0 ? null : new List<string> { Word(random) },
                Optional = random.Next(2) == 0
            };
        }
        if (random.Next(2) == 0) manifest.DevDependencies = new Dictionary<string, Dependency> { ["insta"] = new Dependency { Version = "1" } };
        for (int i = random.Next(3); i > 0; i--) manifest.Features["f" + i] = Enumerable.Range(0, random.Next(3)).Select(_ => Word(random)).ToList();
        for (int i = random.Next(3); i > 0; i--) {
            manifest.Bin.Add(new BinTarget {
                Name = Word(random),
                Path = "src/bin/" + i + ".rs",
                Env = random.Next(2) == 0 ? null : new Dictionary<string, string> { ["RUST_LOG"] = Word(random) }
            });
        }
        if (random.Next(2) == 0) {
            manifest.Profile = new CargoProfile {
                Release = new ProfileSettings { OptLevel = 3, Lto = true, CodegenUnits = random.Next(2) == 0 ? null : 1 },
                Dev = random.Next(2) == 0 ? null : new ProfileSettings()
            };
        }
        return manifest;
    }

    public static Scalars Scalars(Random random) {
        double[] doubles = { 0.0, -0.0, 1.5, 1e300, -2.5e-300, double.PositiveInfinity, double.NegativeInfinity, double.NaN, Math.PI };
        return new Scalars {
            Text = Word(random),
            Letter = (char)('a' + random.Next(26)),
            Flag = random.Next(2) == 0,
            Byte = (byte)random.Next(256),
            SByte = (sbyte)random.Next(-128, 128),
            Short = (short)random.Next(short.MinValue, short.MaxValue),
            UShort = (ushort)random.Next(ushort.MaxValue),
            Int = random.Next(int.MinValue, int.MaxValue),
            UInt = (uint)random.Next() * 2u,
            Long = random.Next(2) == 0 ? long.MinValue : (long)random.Next() << 20,
            ULong = random.Next(2) == 0 ? long.MaxValue : (ulong)random.Next(),
            Big = new BigInteger(random.Next()) * 1000,
            Single = random.Next(2) == 0 ? 1.1f : float.NegativeInfinity,
            Double = doubles[random.Next(doubles.Length)],
            Decimal = random.Next(2) == 0 ? 0.1m : 79228162514264337593543950335m,
            Instant = new DateTimeOffset(2000 + random.Next(30), 1 + random.Next(12), 1 + random.Next(28), random.Next(24), random.Next(60), random.Next(60), TimeSpan.FromMinutes(random.Next(-12, 13) * 30)).AddTicks(random.Next(10_000_000)),
            Local = new DateTime(2000 + random.Next(30), 1 + random.Next(12), 1 + random.Next(28), random.Next(24), random.Next(60), random.Next(60), DateTimeKind.Unspecified),
#if NET8_0_OR_GREATER
            Day = new DateOnly(1979, 5, 27).AddDays(random.Next(10_000)),
            Clock = new TimeOnly(random.Next(24), random.Next(60), random.Next(60), random.Next(1000)),
#endif
            Elapsed = random.Next(2) == 0 ? TimeSpan.FromSeconds(random.Next(86_400)) : TimeSpan.FromDays(-1.5),
            Raw = random.Next(2) == 0 ? TomlDateTime.FromDate(2024, 2, 29) : TomlDateTime.FromTime(7, 32, 0, 999_999_900),
            Id = new Guid(random.Next(), 1, 2, new byte[] { 1, 2, 3, 4, 5, 6, 7, 8 }),
            Link = random.Next(2) == 0 ? null : new Uri("https://example.com/a?b=" + random.Next()),
            Blob = random.Next(2) == 0 ? null : new byte[] { 0, 255, (byte)random.Next(256) },
            Flags = (Features)random.Next(8),
            MaybeEdition = random.Next(2) == 0 ? null : Edition.E2021
        };
    }

    public static Collections Collections(Random random) => new Collections {
        Numbers = Enumerable.Range(0, random.Next(4)).Select(i => i * random.Next(100)).ToArray(),
        Matrix = Enumerable.Range(0, random.Next(3)).Select(_ => Enumerable.Range(0, random.Next(3)).ToList()).ToList(),
        Unique = new HashSet<string>(Enumerable.Range(0, random.Next(3)).Select(i => Word(random) + i)),
        Readings = Enumerable.Range(0, random.Next(3)).Select(i => i / 4.0).ToList(),
        ById = Enumerable.Range(0, random.Next(3)).ToDictionary(i => i - 1, i => Word(random)),
        ByEdition = new Dictionary<Edition, bool> { [Edition.E2018] = random.Next(2) == 0 },
        Points = Enumerable.Range(0, random.Next(3)).Select(i => new Point { X = i, Y = -i }).ToList(),
        Origin = new Point { X = random.Next(9), Y = random.Next(9) },
        Anything = random.Next(3) switch {
            0 => null,
            1 => new Dictionary<string, object?> { ["n"] = 1L, ["list"] = new List<object?> { "a", 2.5 } },
            _ => new List<object?> { new Dictionary<string, object?> { ["k"] = true } }
        }
    };

    public static Team Team(Random random) => new Team {
        Name = Word(random),
        Members = Enumerable.Range(0, random.Next(4)).Select(i => new Person(Word(random), random.Next(90), random.Next(2) == 0 ? null : "p" + i + "@example.com")).ToList()
    };

    public static Monitoring Monitoring(Random random) => new Monitoring {
        Probe = random.Next(3) switch {
            0 => null,
            1 => new HttpProbe { Interval = random.Next(60), Path = "/health" },
            _ => new TcpProbe { Interval = 5, Port = random.Next(65536) }
        },
        Probes = Enumerable.Range(0, random.Next(3)).Select(i => i % 2 == 0 ? (Probe)new HttpProbe { Interval = i } : new TcpProbe { Port = i }).ToList(),
        Shapes = Enumerable.Range(0, random.Next(3)).Select(i => i % 2 == 0 ? new Shape { Color = "red" } : new Circle { Radius = i }).ToList(),
        Window = new Size(random.Next(4000), random.Next(3000))
    };
}
