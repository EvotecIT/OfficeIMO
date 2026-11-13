using Xunit;

namespace OfficeIMO.Toml.Tests.Serialization;

// Contracts of the TOML mapping that a round trip alone cannot show: where tables, inline tables and arrays of tables
// go, how each date-time kind and special number is written, what happens to null, and what errors report.
public class TomlSerializerBehaviorTests {
    private static readonly TomlSerializerOptions Reflection = new TomlSerializerOptions();

    [Fact]
    public void PairsComeBeforeSectionsAndTablesWithOnlySubTablesGetNoHeader() {
        var manifest = new CargoManifest {
            Package = new CargoPackage { Name = "demo", Edition = Edition.E2021 },
            Dependencies = { ["serde"] = new Dependency { Version = "1", Features = new List<string> { "derive" } } },
            Bin = { new BinTarget { Name = "a", Path = "a.rs", Env = new Dictionary<string, string> { ["K"] = "v" } }, new BinTarget { Name = "b", Path = "b.rs" } },
            Profile = new CargoProfile { Release = new ProfileSettings { OptLevel = 3 } }
        };
        var options = new TomlSerializerOptions(TestTomlContext.Default.Options) { PropertyNamingPolicy = OfficeNamingPolicy.KebabCaseLower };
        string toml = TomlSerializer.Serialize(manifest, (TomlTypeInfo<CargoManifest>)options.GetTypeInfo(typeof(CargoManifest)));
        Assert.Equal(
            "[package]\nname = \"demo\"\nversion = \"0.1.0\"\nedition = \"E2021\"\nauthors = []\npublish = true\n" +
            "\n[dependencies]\nserde = { version = \"1\", features = [\"derive\"], optional = false }\n" +
            "\n[features]\n" +
            "\n[[bin]]\nname = \"a\"\npath = \"a.rs\"\n\n[bin.env]\nK = \"v\"\n" +
            "\n[[bin]]\nname = \"b\"\npath = \"b.rs\"\n" +
            "\n[profile.release]\nopt-level = 3\nlto = false\n", toml);
    }

    [Fact]
    public void InlineTableAttributeOnAPropertyWritesItsTablesOnTheKeyLine() {
        var collections = new Collections { Origin = new Point { X = 1, Y = 2 }, Points = { new Point { X = 3, Y = 4 } } };
        string toml = TomlSerializer.Serialize(collections, TestTomlContext.Default.Collections);
        Assert.Contains("\nOrigin = { X = 1, Y = 2 }\n", toml);
        // A collection of tables outside an inline context is an array of tables.
        Assert.Contains("\n[[Points]]\nX = 3\nY = 4\n", toml);
    }

    [Fact]
    public void ArraysOfTablesAndInlineFormsReadBackFromEitherSpelling() {
        const string sections = "[[Members]]\nName = \"Ada\"\nAge = 36\n\n[[Members]]\nName = \"Alan\"\nAge = 41\n";
        const string inline = "Members = [{ Name = \"Ada\", Age = 36 }, { Name = \"Alan\", Age = 41 }]\n";
        var expected = new[] { new Person("Ada", 36), new Person("Alan", 41) };
        Assert.Equal(expected, TomlSerializer.Deserialize(sections, TestTomlContext.Default.Team).Members);
        Assert.Equal(expected, TomlSerializer.Deserialize(inline, TestTomlContext.Default.Team).Members);
    }

    [Fact]
    public void DateTimeKindsMapToTheirTomlKinds() {
        var scalars = new Scalars {
            Instant = new DateTimeOffset(1979, 5, 27, 7, 32, 0, TimeSpan.FromHours(-7)),
            Local = new DateTime(1979, 5, 27, 7, 32, 0, DateTimeKind.Unspecified),
#if NET8_0_OR_GREATER
            Day = new DateOnly(1979, 5, 27),
            Clock = new TimeOnly(7, 32, 0),
#endif
            Elapsed = TimeSpan.FromDays(2),
            Raw = TomlDateTime.FromDate(2024, 2, 29)
        };
        string toml = TomlSerializer.Serialize(scalars, TestTomlContext.Default.Scalars);
        Assert.Contains("\nInstant = 1979-05-27T07:32:00-07:00\n", toml);
        Assert.Contains("\nLocal = 1979-05-27T07:32:00\n", toml);
#if NET8_0_OR_GREATER
        Assert.Contains("\nDay = 1979-05-27\n", toml);
        Assert.Contains("\nClock = 07:32:00\n", toml);
#endif
        Assert.Contains("\nElapsed = \"2.00:00:00\"\n", toml);
        Assert.Contains("\nRaw = 2024-02-29\n", toml);

        Scalars read = TomlSerializer.Deserialize("Local = 1979-05-27T07:32:00Z\nElapsed = 07:32:00\nRaw = 07:32:00", TestTomlContext.Default.Scalars);
        // An offset date-time read as DateTime is the UTC instant; a local time is a time of day.
        Assert.Equal(new DateTime(1979, 5, 27, 7, 32, 0, DateTimeKind.Utc), read.Local);
        Assert.Equal(DateTimeKind.Utc, read.Local.Kind);
        Assert.Equal(new TimeSpan(7, 32, 0), read.Elapsed);
        Assert.Equal(TomlDateTimeKind.LocalTime, read.Raw.Kind);
        // A local date-time has no offset, so it does not read as a DateTimeOffset.
        TomlException error = Assert.Throws<TomlException>(() => TomlSerializer.Deserialize("Instant = 1979-05-27T07:32:00", TestTomlContext.Default.Scalars));
        Assert.Equal(TomlErrorCodes.TypeMismatch, error.Code);
    }

    [Fact]
    public void SpecialAndExactNumbersRoundTrip() {
        var scalars = new Scalars { Double = double.NaN, Single = float.PositiveInfinity, Decimal = 0.1m };
        string toml = TomlSerializer.Serialize(scalars, TestTomlContext.Default.Scalars);
        Assert.Contains("\nDouble = nan\n", toml);
        Assert.Contains("\nSingle = inf\n", toml);
        Assert.Contains("\nDecimal = 0.1\n", toml);
        Scalars read = TomlSerializer.Deserialize("Double = -inf\nDecimal = 1_000.000_000_000_000_000_1\nInt = 3.0\nSingle = 1", TestTomlContext.Default.Scalars);
        Assert.Equal(double.NegativeInfinity, read.Double);
        Assert.Equal(1000.0000000000000001m, read.Decimal);
        Assert.Equal(3, read.Int);
        Assert.Equal(1f, read.Single);
    }

    [Fact]
    public void ValuesOutsideTheTargetRangeFailWithTheirPosition() {
        TomlException tooBig = Assert.Throws<TomlException>(() => TomlSerializer.Deserialize("Text = \"x\"\nByte = 256", TestTomlContext.Default.Scalars));
        Assert.Equal(TomlErrorCodes.TypeMismatch, tooBig.Code);
        Assert.Equal(2, tooBig.Line);
        Assert.Equal(8, tooBig.Column);
        Assert.Contains("Byte", tooBig.Message);

        TomlException unsigned = Assert.Throws<TomlException>(() => TomlSerializer.Serialize(new Scalars { ULong = ulong.MaxValue }, TestTomlContext.Default.Scalars));
        Assert.Equal(TomlErrorCodes.UnrepresentableValue, unsigned.Code);
    }

    [Fact]
    public void TypeMismatchesReportThePathOfTheValue() {
        const string toml = "[tool.ruff]\nline-length = 88\n\n[tool.ruff.lint]\nselect = [\"E\", 4]\n";
        using TomlDocument document = TomlDocument.Parse(toml);
        var options = new TomlSerializerOptions { PropertyNamingPolicy = OfficeNamingPolicy.KebabCaseLower };
        TomlException error = Assert.Throws<TomlException>(() => TomlSerializer.Deserialize<RuffSettings>(document.Select("tool.ruff")!, options));
        Assert.Equal(TomlErrorCodes.TypeMismatch, error.Code);
        Assert.Equal(5, error.Line);
        Assert.Equal(16, error.Column);
        Assert.Contains("tool.ruff.lint.select[1]", error.Message);

        document.Root["tool"].AsTable()["ruff"].AsTable()["lint"].AsTable()["select"].AsArray().RemoveAt(1);
        RuffSettings settings = TomlSerializer.Deserialize<RuffSettings>(document.Select("tool.ruff")!, options);
        Assert.Equal(88, settings.LineLength);
        Assert.Equal(new[] { "E" }, settings.Lint.Select);
    }

    [Fact]
    public void NullsAreLeftOutUnlessTheOptionsAskForAnError() {
        var package = new CargoPackage { Name = "demo", Description = null };
        string toml = TomlSerializer.Serialize(new CargoManifest { Package = package }, TestTomlContext.Default.CargoManifest);
        Assert.DoesNotContain("Description", toml);
        Assert.Null(TomlSerializer.Deserialize(toml, TestTomlContext.Default.CargoManifest).Package.Description);

        var strict = new TestTomlContext(new TomlSerializerOptions { NullHandling = TomlNullHandling.Error });
        TomlException error = Assert.Throws<TomlException>(() => TomlSerializer.Serialize(new CargoManifest { Package = package }, strict.CargoManifest));
        Assert.Equal(TomlErrorCodes.UnrepresentableValue, error.Code);

        // A null item cannot be left out of an array without shifting the items after it.
        var collections = new Collections { Anything = new List<object?> { 1L, null } };
        Assert.Equal(TomlErrorCodes.UnrepresentableValue, Assert.Throws<TomlException>(() => TomlSerializer.Serialize(collections, TestTomlContext.Default.Collections)).Code);
    }

    [Fact]
    public void ConstructorsInitOnlyAndRequiredMembersBind() {
        Account account = TomlSerializer.Deserialize("Id = \"a1\"\nOwner = \"ada\"\nBalance = 10.5\n", TestTomlContext.Default.Account);
        Assert.Equal("a1", account.Id);
        Assert.Equal("ada", account.Owner);
        Assert.Equal(10.5m, account.Balance);

        TomlException missing = Assert.Throws<TomlException>(() => TomlSerializer.Deserialize("Id = \"a1\"\n", TestTomlContext.Default.Account));
        Assert.Equal(TomlErrorCodes.MissingRequiredMember, missing.Code);
        Assert.Contains("Balance", missing.Message);

        Monitoring monitoring = TomlSerializer.Deserialize("[Window]\nHeight = 2\nWidth = 3\n", TestTomlContext.Default.Monitoring);
        Assert.Equal(3, monitoring.Window.Width);
        Assert.Equal(2, monitoring.Window.Height);
    }

    [Fact]
    public void DerivedTypesAreChosenOnlyByRegisteredDiscriminators() {
        Monitoring monitoring = TomlSerializer.Deserialize("[Probe]\nPort = 22\n\"$type\" = \"tcp\"\n\n[[Shapes]]\nkind = \"circle\"\nRadius = 2.0\n\n[[Shapes]]\nColor = \"red\"\n",
            TestTomlContext.Default.Monitoring);
        Assert.Equal(22, Assert.IsType<TcpProbe>(monitoring.Probe).Port);
        Assert.Equal(2.0, Assert.IsType<Circle>(monitoring.Shapes[0]).Radius);
        Assert.Equal("red", Assert.IsType<Shape>(monitoring.Shapes[1]).Color);

        TomlException unknown = Assert.Throws<TomlException>(() => TomlSerializer.Deserialize("[Probe]\n\"$type\" = \"udp\"\n", TestTomlContext.Default.Monitoring));
        Assert.Equal(TomlErrorCodes.UnknownDiscriminator, unknown.Code);
        // The abstract base cannot be created, so a table without a discriminator is an error too.
        Assert.Equal(TomlErrorCodes.UnknownDiscriminator,
            Assert.Throws<TomlException>(() => TomlSerializer.Deserialize("[Probe]\nInterval = 1\n", TestTomlContext.Default.Monitoring)).Code);
    }

    [Fact]
    public void DictionaryKeysReadAsTheirKeyType() {
        Collections collections = TomlSerializer.Deserialize("[ById]\n-1 = \"minus\"\n7 = \"seven\"\n\n[ByEdition]\nE2024 = true\n", TestTomlContext.Default.Collections);
        Assert.Equal("minus", collections.ById[-1]);
        Assert.Equal("seven", collections.ById[7]);
        Assert.True(collections.ByEdition[Edition.E2024]);
        TomlException error = Assert.Throws<TomlException>(() => TomlSerializer.Deserialize("[ById]\nx = \"y\"\n", TestTomlContext.Default.Collections));
        Assert.Equal(TomlErrorCodes.TypeMismatch, error.Code);
    }

    [Fact]
    public void EnumsReadNamesNumbersAndFlags() {
        Scalars read = TomlSerializer.Deserialize("Flags = \"std, serde\"\nMaybeEdition = 3\n", TestTomlContext.Default.Scalars);
        Assert.Equal(Features.Std | Features.Serde, read.Flags);
        Assert.Equal(Edition.E2024, read.MaybeEdition);
        Assert.Equal(Features.Default | Features.Serde, TomlSerializer.Deserialize("Flags = [\"Default\", \"Serde\"]", TestTomlContext.Default.Scalars).Flags);
        Assert.Contains("\nFlags = \"Default\"\n", TomlSerializer.Serialize(new Scalars { Flags = Features.Std | Features.Alloc }, TestTomlContext.Default.Scalars));
    }

    [Fact]
    public void UntypedValuesKeepTheirTomlTypes() {
        Dictionary<string, object> values = TomlSerializer.Deserialize(
            "a = 1\nb = 1.5\nc = \"s\"\nd = 1979-05-27\ne = [1, \"x\"]\n[t]\nk = true\n[[arr]]\nn = 1\n", TestTomlContext.Default.DictionaryStringObject);
        Assert.Equal(1L, values["a"]);
        Assert.Equal(1.5, values["b"]);
        Assert.Equal("s", values["c"]);
        Assert.Equal(TomlDateTime.FromDate(1979, 5, 27), values["d"]);
        Assert.Equal(new List<object?> { 1L, "x" }, values["e"]);
        Assert.Equal(true, ((Dictionary<string, object?>)values["t"])["k"]);
        Assert.Equal(1L, ((Dictionary<string, object?>)((List<object?>)values["arr"])[0])["n"]);
        Assert.Equal("a = 1\nb = 1.5\nc = \"s\"\nd = 1979-05-27\ne = [1, \"x\"]\n\n[t]\nk = true\n\n[[arr]]\nn = 1\n",
            TomlSerializer.Serialize(values, TestTomlContext.Default.DictionaryStringObject));
    }

    [Fact]
    public void WritingFailsForNonTablesAndCycles() {
        Assert.Throws<NotSupportedException>(() => TomlSerializer.Serialize(new List<int> { 1 }, Reflection));
        var node = new Node { Name = "a" };
        node.Child = new Node { Name = "b", Child = node };
        Assert.Equal(TomlErrorCodes.ReferenceCycle, Assert.Throws<TomlException>(() => TomlSerializer.Serialize(node, TestTomlContext.Default.Node)).Code);
    }

    [Fact]
    public void StreamsAndTextReadersCarryTheSameDocument() {
        var team = new Team { Name = "core", Members = { new Person("Ada", 36, "ada@example.com") } };
        using var stream = new MemoryStream();
        TomlSerializer.Serialize(stream, team, TestTomlContext.Default.Team);
        Assert.Equal(TomlSerializer.SerializeToUtf8Bytes(team, TestTomlContext.Default.Team), stream.ToArray());
        stream.Position = 0;
        Assert.Equal(team.Members, TomlSerializer.Deserialize(stream, TestTomlContext.Default.Team).Members);
        var text = new StringWriter();
        TomlSerializer.Serialize(text, team, typeof(Team), TestTomlContext.Default);
        Assert.Equal(team.Members, ((Team)TomlSerializer.Deserialize(new StringReader(text.ToString()), typeof(Team), TestTomlContext.Default)).Members);
    }
}
