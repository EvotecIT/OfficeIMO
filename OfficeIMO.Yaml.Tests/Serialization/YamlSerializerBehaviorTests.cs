using System.Numerics;
using Xunit;

namespace OfficeIMO.Yaml.Tests.Serialization;

public class YamlSerializerBehaviorTests {
    private static TestYamlContext Context => TestYamlContext.Default;

    // Aliases are read as their anchored node, and merge keys add only the entries a mapping does not set itself.
    [Fact]
    public void AliasesExpandAndMergeKeysFillUnsetMembers() {
        const string yaml = "Base: &base\n  Name: shared\n  Port: 80\n  Tags: &tags [a, b]\n"
            + "Configs:\n- <<: *base\n  Port: 8080\n- Port: 9090\n  <<: [*base, {Enabled: true, Port: 1}]\n  Tags: *tags\n- *base\n";
        ConfigSet set = YamlSerializer.Deserialize(yaml, Context.ConfigSet)!;
        AppConfig first = set.Configs[0], second = set.Configs[1], third = set.Configs[2];
        Assert.Equal(("shared", 8080, "a,b"), (first.Name, first.Port, string.Join(",", first.Tags)));
        Assert.Equal(("shared", 9090, true, "a,b"), (second.Name, second.Port, second.Enabled, string.Join(",", second.Tags)));
        Assert.Equal(80, third.Port);
        // Without Preserve an alias is a copy.
        Assert.NotSame(set.Base, third);

        var untyped = (Dictionary<string, object?>)((List<object?>)YamlSerializer.Deserialize<Dictionary<string, object>>(yaml, Context.DictionaryStringObject)!["Configs"]!)[1]!;
        Assert.Equal((9090L, "shared", true), ((long)untyped["Port"]!, (string)untyped["Name"]!, (bool)untyped["Enabled"]!));
    }

    // A "billion laughs" input stops at MaxAliasExpansion instead of exhausting memory.
    [Fact]
    public void AliasExpansionStopsAtTheReaderLimit() {
        var yaml = new System.Text.StringBuilder("a: &a [x, x, x, x, x, x, x, x, x, x]\n");
        for (char level = 'b'; level <= 'h'; level++) {
            yaml.Append(level).Append(": &").Append(level).Append(" [");
            for (int i = 0; i < 10; i++) yaml.Append(i == 0 ? "" : ", ").Append('*').Append((char)(level - 1));
            yaml.Append("]\n");
        }
        var options = new YamlSerializerOptions { ReaderOptions = YamlReaderOptions.UntrustedDefaults };
        var exception = Assert.Throws<YamlLimitExceededException>(() => YamlSerializer.Deserialize(yaml.ToString(), typeof(object), new TestYamlContext(options)));
        Assert.Equal(nameof(YamlReaderOptions.MaxAliasExpansion), exception.LimitName);
    }

    [Fact]
    public void AnAliasInsideItsOwnAnchorFailsWithoutPreserve() {
        var exception = Assert.Throws<YamlException>(() => YamlSerializer.Deserialize("&a [1, *a]", typeof(object), Context));
        Assert.Equal(YamlErrorCodes.RecursiveAlias, exception.Code);
    }

    // Preserve writes one anchor per shared object and aliases for later references, and reads them back as one instance.
    [Fact]
    public void PreserveRoundTripsSharedReferencesAndCycles() {
        var context = new TestYamlContext(new YamlSerializerOptions { ReferenceHandling = YamlReferenceHandling.Preserve });
        var root = new TreeNode { Name = "root" };
        var child = new TreeNode { Name = "child", Parent = root };
        root.Children.Add(child);
        root.Children.Add(child);

        string yaml = YamlSerializer.Serialize(root, context.TreeNode);
        Assert.Equal("&id001\nName: root\nParent: null\nChildren:\n- &id002\n  Name: child\n  Parent: *id001\n  Children: []\n- *id002\n", yaml);

        TreeNode copy = YamlSerializer.Deserialize(yaml, context.TreeNode)!;
        Assert.Same(copy, copy.Children[0].Parent);
        Assert.Same(copy.Children[0], copy.Children[1]);
    }

    [Fact]
    public void CyclesFailWithoutPreserve() {
        var root = new TreeNode { Name = "root" };
        root.Children.Add(root);
        var exception = Assert.Throws<YamlException>(() => YamlSerializer.Serialize(root, Context.TreeNode));
        Assert.Equal(YamlErrorCodes.ReferenceCycle, exception.Code);
    }

    // Input that does not fit the model fails at the node, with its line, column and path.
    [Theory]
    [InlineData("Name: app\nPort: eighty\n", 2, 7, "Port")]
    [InlineData("Name: app\nTags:\n  key: value\n", 3, 3, "Tags")]
    [InlineData("Name: app\nLabels:\n  a: [1, 2]\n", 3, 6, "Labels.a")]
    [InlineData("Name: app\nPort: 99999999999\n", 2, 7, "Port")]
    [InlineData("Port: ~\n", 1, 7, "Port")]
    public void MismatchesFailAtTheNodeWithItsPath(string yaml, int line, int column, string path) {
        var exception = Assert.Throws<YamlException>(() => YamlSerializer.Deserialize(yaml, Context.AppConfig));
        Assert.Equal(YamlErrorCodes.TypeMismatch, exception.Code);
        Assert.Equal((line, column), (exception.Line, exception.Column));
        Assert.Contains(" at " + path, exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RequiredMembersMustBePresent() {
        var exception = Assert.Throws<YamlException>(() => YamlSerializer.Deserialize("Owner: me\nBalance: 1\n", Context.Account));
        Assert.Equal(YamlErrorCodes.MissingRequiredMember, exception.Code);
        Assert.Contains("'Id'", exception.Message, StringComparison.Ordinal);
        exception = Assert.Throws<YamlException>(() => YamlSerializer.Deserialize("Id: a1\n", Context.Account));
        Assert.Contains("'Balance'", exception.Message, StringComparison.Ordinal);
        Account account = YamlSerializer.Deserialize("Id: a1\nBalance: 2.50\n", Context.Account)!;
        // An init-only member the mapping leaves out keeps its initializer.
        Assert.Equal(("a1", "unknown", 2.50m), (account.Id, account.Owner, account.Balance));
    }

    [Fact]
    public void DuplicateKeysTakeTheLastValueUnlessStrict() {
        Assert.Equal("b", YamlSerializer.Deserialize("Name: a\nName: b\n", Context.AppConfig)!.Name);
        Assert.Equal(2L, YamlSerializer.Deserialize("a: 1\na: 2\n", Context.DictionaryStringObject)!["a"]);

        var strict = new TestYamlContext(new YamlSerializerOptions { ReaderOptions = new YamlReaderOptions { Strict = true } });
        Assert.Equal(YamlErrorCodes.DuplicateKey, Assert.Throws<YamlException>(() => YamlSerializer.Deserialize("Name: a\nName: b\n", strict.AppConfig)).Code);
        Assert.Equal(YamlErrorCodes.DuplicateKey, Assert.Throws<YamlException>(() => YamlSerializer.Deserialize("a: 1\na: 2\n", strict.DictionaryStringObject)).Code);
    }

    // Only registered discriminators create derived types; a tag never names a .NET type by itself.
    [Fact]
    public void PolymorphismAcceptsOnlyRegisteredDiscriminators() {
        Zoo zoo = YamlSerializer.Deserialize("Animals:\n- !dog {Name: Rex, GoodBoy: true}\n- !cat\n  Name: Tom\n  Lives: 9\n", Context.Zoo)!;
        Assert.IsType<Dog>(zoo.Animals[0]);
        Assert.Equal(9, Assert.IsType<Cat>(zoo.Animals[1]).Lives);

        YamlException unknown = Assert.Throws<YamlException>(() => YamlSerializer.Deserialize("Animals:\n- !System.IO.FileInfo {Name: x}\n", Context.Zoo));
        Assert.Equal(YamlErrorCodes.UnknownDiscriminator, unknown.Code);
        Assert.Equal((2, 3), (unknown.Line, unknown.Column));
        YamlException untagged = Assert.Throws<YamlException>(() => YamlSerializer.Deserialize("Animals:\n- {Name: x}\n", Context.Zoo));
        Assert.Equal(YamlErrorCodes.UnknownDiscriminator, untagged.Code);

        object? anything = YamlSerializer.Deserialize<Collections>("Anything: !System.Diagnostics.Process {StartInfo: x}\n", Context.Collections)!.Anything;
        Assert.Equal(new Dictionary<string, object?> { ["StartInfo"] = "x" }, anything);
    }

    // A property discriminator may come anywhere in the mapping; it is written first.
    [Theory]
    [InlineData("kind: Service\napiVersion: v1\nmetadata: {name: web}\nspec: {ports: [{port: 80}]}\n")]
    [InlineData("apiVersion: v1\nkind: Service\nmetadata: {name: web}\nspec: {ports: [{port: 80}]}\n")]
    [InlineData("apiVersion: v1\nmetadata: {name: web}\nspec: {ports: [{port: 80}]}\nkind: Service\n")]
    public void PropertyDiscriminatorsAreFoundAnywhere(string yaml) {
        var context = new TestYamlContext(new YamlSerializerOptions { PropertyNamingPolicy = YamlNamingPolicy.CamelCase, DefaultIgnoreCondition = YamlIgnoreCondition.WhenWritingNull });
        KubeService service = Assert.IsType<KubeService>(YamlSerializer.Deserialize(yaml, context.KubeObject));
        Assert.Equal(("web", 80), (service.Metadata.Name, service.Spec.Ports[0].Port));
        Assert.Equal("kind: Service\napiVersion: v1\nmetadata:\n  name: web\nspec:\n  ports:\n  - port: 80\n", YamlSerializer.Serialize<KubeObject>(service, context.KubeObject));
    }

    // Flags enums are one comma-separated scalar, as Enum.ToString writes them; sequences and numbers also read.
    [Fact]
    public void EnumsAreNamesAndFlagsAreCommaSeparated() {
        var config = new AppConfig { Level = LogLevel.Warning, Access = Permissions.Read | Permissions.Execute };
        string yaml = YamlSerializer.Serialize(config, Context.AppConfig);
        Assert.Contains("Level: Warning\n", yaml, StringComparison.Ordinal);
        Assert.Contains("Access: Read, Execute\n", yaml, StringComparison.Ordinal);
        Assert.Contains("Access: ReadWrite\n", YamlSerializer.Serialize(new AppConfig { Access = Permissions.ReadWrite }, Context.AppConfig), StringComparison.Ordinal);
        Assert.Contains("Access: 16\n", YamlSerializer.Serialize(new AppConfig { Access = (Permissions)16 }, Context.AppConfig), StringComparison.Ordinal);

        foreach (string access in new[] { "read, execute", "[Read, Execute]", "5" }) {
            Assert.Equal(Permissions.Read | Permissions.Execute, YamlSerializer.Deserialize("Access: " + access + "\n", Context.AppConfig)!.Access);
        }
        Assert.Equal(YamlErrorCodes.TypeMismatch, Assert.Throws<YamlException>(() => YamlSerializer.Deserialize("Level: Loud\n", Context.AppConfig)).Code);
    }

    // The documented scalar formats: ISO 8601 timestamps written plain, time spans and times quoted only where YAML 1.1
    // reads them as sexagesimal numbers (12:30:00 but not 08:30:00), shortest round-trip floats, !!binary base64.
    [Fact]
    public void ScalarsUseTheDocumentedFormats() {
        var scalars = new Scalars {
            Text = "on",
            Letter = 'y',
            ULong = ulong.MaxValue,
            Single = 0.1f,
            Double = 1e21,
            Decimal = 1.50m,
            Big = BigInteger.Pow(10, 30),
            When = new DateTime(2024, 5, 1, 8, 30, 0, DateTimeKind.Utc),
            At = new DateTimeOffset(2024, 5, 1, 8, 30, 0, 500, TimeSpan.FromHours(2)),
            Duration = new TimeSpan(0, 12, 30, 0),
            Id = new Guid("0f8fad5b-d9cb-469f-a165-70867728950e"),
            Payload = new byte[] { 1, 2, 3 },
#if NET6_0_OR_GREATER
            Day = new DateOnly(2024, 5, 1),
            Clock = new TimeOnly(8, 30),
#endif
        };
        string yaml = YamlSerializer.Serialize(scalars, Context.Scalars);
        Assert.Contains("Text: 'on'\nLetter: 'y'\n", yaml, StringComparison.Ordinal);
        Assert.Contains("ULong: 18446744073709551615\nSingle: 0.1\nDouble: 1.0e+21\nDecimal: 1.50\nBig: 1000000000000000000000000000000\n", yaml, StringComparison.Ordinal);
        Assert.Contains("When: 2024-05-01T08:30:00Z\nAt: 2024-05-01T08:30:00.5+02:00\nDuration: '12:30:00'\nId: 0f8fad5b-d9cb-469f-a165-70867728950e\nPayload: !!binary AQID\n", yaml, StringComparison.Ordinal);
#if NET6_0_OR_GREATER
        Assert.Contains("Day: 2024-05-01\nClock: 08:30:00\n", yaml, StringComparison.Ordinal);
#endif
        // YAML 1.1 spellings read too.
        Scalars read = YamlSerializer.Deserialize("When: 2001-12-14 21:59:43.10Z\nDuration: 00:30:00\nBig: 0x10\nPayload: !!binary |\n  AQ\n  ID\n", Context.Scalars)!;
        Assert.Equal(new DateTime(2001, 12, 14, 21, 59, 43, 100, DateTimeKind.Utc), read.When.ToUniversalTime());
        Assert.Equal(TimeSpan.FromMinutes(30), read.Duration);
        Assert.Equal(new BigInteger(16), read.Big);
        Assert.Equal(new byte[] { 1, 2, 3 }, read.Payload);
    }

    // Values are read by their target type, whatever their quoting; a quoted 'null' stays text.
    [Fact]
    public void ScalarsAreReadByTargetType() {
        AppConfig config = YamlSerializer.Deserialize("Name: 'null'\nPort: \"8080\"\nEnabled: 'true'\nRatio: 1\nComment:\nRetries: ~\n", Context.AppConfig)!;
        Assert.Equal(("null", 8080, true, 1.0, null, null), (config.Name, config.Port, config.Enabled, config.Ratio, config.Comment, config.Retries));
        var yaml11 = YamlSerializer.Deserialize("%YAML 1.1\n---\nEnabled: yes\nPort: 0x1F\n", Context.AppConfig)!;
        Assert.Equal((true, 31), (yaml11.Enabled, yaml11.Port));
    }

    // Naming policies, [YamlPropertyName], [YamlIgnore] and the ignore conditions decide which keys are written.
    [Fact]
    public void NamingAndIgnoreConditionsShapeTheKeys() {
        var config = new AppConfig { Name = "app", Comment = null, Retries = 0 };
        var snake = new TestYamlContext(new YamlSerializerOptions { PropertyNamingPolicy = YamlNamingPolicy.SnakeCaseLower, DefaultIgnoreCondition = YamlIgnoreCondition.WhenWritingDefault });
        string yaml = YamlSerializer.Serialize(config, snake.AppConfig);
        Assert.Equal("name: app\nretries: 0\ntags: []\nlabels: {}\nhosts: []\nfeature-flags: {}\n", yaml);
        Assert.DoesNotContain("Secret", YamlSerializer.Serialize(config, Context.AppConfig), StringComparison.Ordinal);
        Assert.Equal("never written", YamlSerializer.Deserialize("Secret: read\n", Context.AppConfig)!.Secret);

        Assert.Equal("imagePullPolicy", YamlNamingPolicy.CamelCase.ConvertName("ImagePullPolicy"));
        Assert.Equal("urlValue", YamlNamingPolicy.CamelCase.ConvertName("URLValue"));
        Assert.Equal("http_server_v2", YamlNamingPolicy.SnakeCaseLower.ConvertName("HTTPServerV2"));
        Assert.Equal("timeout-minutes", YamlNamingPolicy.KebabCaseLower.ConvertName("TimeoutMinutes"));
    }

    // Untyped reads keep file order and resolve scalars through the document's schema.
    [Fact]
    public void UntypedGraphsKeepFileOrderAndSchemaTypes() {
        const string yaml = "zeta: 1\nalpha: [true, 2.5, ~, '3', 99999999999999999999999]\n200: ok\nnested: {b: 1, a: 2}\nwhen: 2001-12-14\n";
        var root = (Dictionary<string, object?>)YamlSerializer.Deserialize(yaml, typeof(object), Context)!;
        Assert.Equal(new[] { "zeta", "alpha", "200", "nested", "when" }, root.Keys);
        Assert.Equal(1L, root["zeta"]);
        Assert.Equal(new object?[] { true, 2.5, null, "3", BigInteger.Parse("99999999999999999999999") }, (List<object?>)root["alpha"]!);
        Assert.Equal(new[] { "b", "a" }, ((Dictionary<string, object?>)root["nested"]!).Keys);
        Assert.Equal("2001-12-14", root["when"]);

        var yaml11 = (Dictionary<string, object?>)YamlSerializer.Deserialize("%YAML 1.1\n---\nflag: on\noctal: 010\n", typeof(object), Context)!;
        Assert.Equal((true, 8L), ((bool)yaml11["flag"]!, (long)yaml11["octal"]!));
    }

    // Writing object-typed values uses the runtime type: dictionaries, lists and generated types.
    [Fact]
    public void ObjectValuesAreWrittenByRuntimeType() {
        var value = new Dictionary<string, object?> {
            ["point"] = new Point { X = 1, Y = 2 },
            ["list"] = new[] { 1, 2 },
            ["level"] = LogLevel.Error,
            ["map"] = new System.Collections.Hashtable { ["k"] = "v" }
        };
        Assert.Equal("point:\n  X: 1\n  'Y': 2\nlist:\n- 1\n- 2\nlevel: Error\nmap:\n  k: v\n", YamlSerializer.Serialize(value, Context.DictionaryStringObject));
        var missing = Assert.Throws<NotSupportedException>(() => YamlSerializer.Serialize(new Dictionary<string, object?> { ["x"] = new Version(1, 0) }, Context.DictionaryStringObject));
        Assert.Contains("[YamlSerializable(typeof(Version))]", missing.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void StreamTextAndReaderOverloadsAgree() {
        var config = new AppConfig { Name = "app", Tags = { "a" } };
        string expected = YamlSerializer.Serialize(config, Context.AppConfig);

        using var stream = new MemoryStream();
        YamlSerializer.Serialize(stream, config, Context.AppConfig);
        Assert.Equal(expected, System.Text.Encoding.UTF8.GetString(stream.ToArray()));
        stream.Position = 0;
        Assert.Equal(expected, YamlSerializer.Serialize(YamlSerializer.Deserialize(stream, Context.AppConfig), Context.AppConfig));
        Assert.Equal(expected, YamlSerializer.Serialize(YamlSerializer.Deserialize(new StringReader(expected), Context.AppConfig), Context.AppConfig));

        // A value written as one node inside a document the caller writes.
        var text = new StringWriter();
        using (var writer = new YamlWriter(text)) {
            writer.WriteStartMapping();
            writer.WriteKey("config");
            YamlSerializer.Serialize(writer, config, Context.AppConfig);
            writer.WriteEndMapping();
        }
        Assert.StartsWith("config:\n  Name: app\n", text.ToString(), StringComparison.Ordinal);

        Assert.Null(YamlSerializer.Deserialize("", Context.AppConfig));
        Assert.Equal(YamlErrorCodes.InvalidDocument, Assert.Throws<YamlException>(() => YamlSerializer.Deserialize("Name: a\n---\nName: b\n", Context.AppConfig)).Code);
    }

    // Typed sections read from and written into a loaded document; untouched text, comments included, saves unchanged.
    [Fact]
    public void TypedSectionsMixWithLosslessDocumentEdits() {
        const string yaml = "# service settings\nname: web   # the app\nspec:\n  Name: app\n  Port: 80\nother: keep   # aligned\n";
        YamlDocument document = YamlDocument.Parse(yaml);
        AppConfig spec = YamlSerializer.Deserialize(document.Select("spec")!, Context.AppConfig)!;
        Assert.Equal(("app", 80), (spec.Name, spec.Port));

        // Leave out what the section did not have: defaults and the collections set to null.
        var compact = new TestYamlContext(new YamlSerializerOptions { DefaultIgnoreCondition = YamlIgnoreCondition.WhenWritingDefault });
        spec.Port = 8080;
        spec.Tags = null!;
        spec.Labels = null!;
        spec.Hosts = null!;
        spec.Features = null!;
        document.Set("spec", YamlSerializer.SerializeToNode(spec, compact.AppConfig));
        Assert.Equal("# service settings\nname: web   # the app\nspec:\n  Name: app\n  Port: 8080\nother: keep   # aligned\n", document.ToString());

        YamlNode node = YamlSerializer.SerializeToNode(new Person("Ada", 36), Context.Person);
        Assert.Equal(new Person("Ada", 36), YamlSerializer.Deserialize(node, Context.Person));

        // Errors inside a loaded node point into the document's text.
        YamlDocument invalid = YamlDocument.Parse("a: 1\nspec:\n  Port: x\n");
        YamlException exception = Assert.Throws<YamlException>(() => YamlSerializer.Deserialize(invalid.Select("spec")!, Context.AppConfig));
        Assert.Equal((YamlErrorCodes.TypeMismatch, 3, 9), (exception.Code, exception.Line, exception.Column));
    }

    [Fact]
    public void MetadataSourcesFailClearly() {
        var options = new YamlSerializerOptions { TypeInfoResolver = Context };
        _ = options.GetTypeInfo(typeof(AppConfig));
        Assert.Throws<InvalidOperationException>(() => options.PropertyNamingPolicy = YamlNamingPolicy.CamelCase);
        var missing = Assert.Throws<NotSupportedException>(() => new YamlSerializerOptions().GetTypeInfo(typeof(AppConfig)));
        Assert.Contains("YamlSerializerContext", missing.Message, StringComparison.Ordinal);
        Assert.Null(Context.GetTypeInfo(typeof(Version)));
    }
    // Properties with a plain name, string dictionaries and string lists write string, integer and boolean entries and
    // items in one step. The text must be what the writer produces for the keys and values written one by one, at the
    // root, nested and as a compact sequence item.
    [Fact]
    public void OneStepEntriesMatchKeyThenValueWrites() {
        string[] texts = {
            "", "plain", "yes", "8080", "100m", "128Mi", "a: b", "#x", "x #y", " lead", "trail ", "multi\nline", "it's", "ünï",
            new string('x', 1200), "-", "- item", "~", "null", "1.5", "+_1.", "registry.example.com/app:1.2.0", "[x]", "{x}"
        };
        var options = new YamlSerializerOptions { PropertyNamingPolicy = YamlNamingPolicy.CamelCase };
        var root = new List<EntryShapes>();
        for (int i = 0; i < texts.Length; i++) {
            root.Add(new EntryShapes {
                Text = texts[i], Number = i * 7 - 50, Name = texts[(i + 5) % texts.Length], Big = long.MinValue + i, Flag = i % 2 == 0,
                Inner = new EntryLeaf { Value = texts[(i + 11) % texts.Length], Count = i },
                Labels = { [texts[i]] = texts[(i + 3) % texts.Length], ["k" + i] = texts[(i + 7) % texts.Length] },
                Tags = { texts[i], texts[(i + 13) % texts.Length] },
                Items = { new EntryShapes { Text = texts[texts.Length - 1 - i], Number = i, Flag = i % 3 == 0 } }
            });
        }
        var single = new StringWriter();
        using (var writer = new YamlWriter(single)) Write(writer, root[3]);
        Assert.Equal(single.ToString(), YamlSerializer.Serialize(root[3], options));

        var expected = new StringWriter();
        using (var writer = new YamlWriter(expected)) {
            writer.WriteStartSequence();
            foreach (EntryShapes item in root) Write(writer, item);
            writer.WriteEndSequence();
        }
        Assert.Equal(expected.ToString(), YamlSerializer.Serialize(root, options));

        static void Write(YamlWriter writer, EntryShapes value) {
            writer.WriteStartMapping();
            writer.WriteKey("text");
            writer.WriteScalar(value.Text);
            writer.WriteKey("number");
            writer.WriteScalar(value.Number);
            writer.WriteKey("name");
            writer.WriteScalar(value.Name);
            writer.WriteKey("big");
            writer.WriteScalar(value.Big);
            writer.WriteKey("flag");
            writer.WriteScalar(value.Flag);
            writer.WriteKey("inner");
            writer.WriteStartMapping();
            writer.WriteKey("value");
            writer.WriteScalar(value.Inner.Value);
            writer.WriteKey("count");
            writer.WriteScalar(value.Inner.Count);
            writer.WriteEndMapping();
            writer.WriteKey("labels");
            writer.WriteStartMapping();
            foreach (KeyValuePair<string, string> label in value.Labels) {
                writer.WriteKey(label.Key);
                writer.WriteScalar(label.Value);
            }
            writer.WriteEndMapping();
            writer.WriteKey("tags");
            writer.WriteStartSequence();
            foreach (string tag in value.Tags) writer.WriteScalar(tag);
            writer.WriteEndSequence();
            writer.WriteKey("items");
            writer.WriteStartSequence();
            foreach (EntryShapes item in value.Items) Write(writer, item);
            writer.WriteEndSequence();
            writer.WriteEndMapping();
        }
    }

    // String and number properties read live scalars without their converters' dispatch unless references are
    // preserved. Both routes must give the same value, or fail the same way, for every spelling in both schemas.
    [Fact]
    public void DirectScalarReadsMatchConverterReads() {
        string[] spellings = {
            "1", "-1", "+1", "0x1F", "0o17", "017", "0b101", "1_000", "1:20", "2147483648", "-2147483649", "9223372036854775808",
            "1.5", "1.50", ".5", "1.", "1e3", "1.0e+3", "6.8523e-05", ".inf", "-.Inf", ".nan", "true", "True", "yes", "on", "off",
            "~", "null", "", "'5'", "\"5\"", "!!str 5", "!!int 5", "!custom 3", "abc", "1,000", "0.0", "-0.0", "&a 7", "|\n  12\n"
        };
        var direct = new YamlSerializerOptions { PropertyNamingPolicy = YamlNamingPolicy.CamelCase };
        var converters = new YamlSerializerOptions { PropertyNamingPolicy = YamlNamingPolicy.CamelCase, ReferenceHandling = YamlReferenceHandling.Preserve };
        foreach (string header in new[] { "", "%YAML 1.1\n---\n" }) {
            foreach (string property in new[] { "text", "number", "big", "flag", "ratio", "amount" }) {
                foreach (string spelling in spellings) {
                    string yaml = header + property + ": " + spelling + "\n";
                    Assert.True(Read(yaml, direct) == Read(yaml, converters), yaml);
                }
            }
        }

        static string Read(string yaml, YamlSerializerOptions options) {
            try {
                ScalarShapes value = YamlSerializer.Deserialize<ScalarShapes>(yaml, options)!;
                return (value.Text ?? "<null>") + "|" + value.Number + "|" + value.Big + "|" + value.Flag + "|"
                    + BitConverter.DoubleToInt64Bits(value.Ratio) + "|" + string.Join(",", decimal.GetBits(value.Amount));
            } catch (YamlException exception) {
                return exception.GetType().Name + ": " + exception.Message;
            }
        }
    }

    public sealed class ScalarShapes {
        public string? Text { get; set; } = "unset";
        public int Number { get; set; }
        public long Big { get; set; }
        public bool Flag { get; set; }
        public double Ratio { get; set; }
        public decimal Amount { get; set; }
    }

    public sealed class EntryShapes {
        public string Text { get; set; } = string.Empty;
        public int Number { get; set; }
        public string Name { get; set; } = string.Empty;
        public long Big { get; set; }
        public bool Flag { get; set; }
        public EntryLeaf Inner { get; set; } = new EntryLeaf();
        public Dictionary<string, string> Labels { get; set; } = new Dictionary<string, string>();
        public List<string> Tags { get; set; } = new List<string>();
        public List<EntryShapes> Items { get; set; } = new List<EntryShapes>();
    }

    public sealed class EntryLeaf {
        public string Value { get; set; } = string.Empty;
        public int Count { get; set; }
    }
}
