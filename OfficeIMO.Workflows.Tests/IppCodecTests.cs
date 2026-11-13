namespace OfficeIMO.Workflows.Tests;

public sealed class IppCodecTests {
    [Fact]
    public void DecodesAHandEncodedRfc8010Response() {
        // IPP/1.1 successful-ok, request-id 7, operation group with charset and language,
        // printer group with a 1setOf keyword and a nameWithLanguage value.
        byte[] bytes = Concat(
            new byte[] { 0x01, 0x01, 0x00, 0x00, 0x00, 0x00, 0x00, 0x07, 0x01 },
            Value(0x47, "attributes-charset", Ascii("utf-8")),
            Value(0x48, "attributes-natural-language", Ascii("en")),
            new byte[] { 0x04 },
            Value(0x44, "sides-supported", Ascii("one-sided")),
            Value(0x44, "", Ascii("two-sided-long-edge")),
            Value(0x36, "printer-name", Concat(new byte[] { 0x00, 0x05 }, Ascii("pl-pl"), new byte[] { 0x00, 0x04 }, Ascii("Druk"))),
            new byte[] { 0x03 },
            Ascii("%PDF"));

        IppMessage message = IppCodec.Decode(bytes, bytes.Length, null, out int dataOffset);

        Assert.Equal(1, message.VersionMajor);
        Assert.Equal(1, message.VersionMinor);
        Assert.Equal(0, message.Code);
        Assert.Equal(7, message.RequestId);
        Assert.Equal(bytes.Length - 4, dataOffset);
        IppAttributeGroup printer = message.FirstGroup(IppGroupTag.Printer)!;
        Assert.Equal(new[] { "one-sided", "two-sided-long-edge" }, printer["sides-supported"]!.Values.Select(v => v.Text));
        IppValue name = printer["printer-name"]!.First!;
        Assert.Equal(IppValueTag.NameWithLanguage, name.Tag);
        Assert.Equal("Druk", name.Text);
        Assert.Equal("pl-pl", name.Language);
    }

    [Fact]
    public void EncodesARequestByteForByte() {
        var request = new IppMessage(2, 0, (short)IppOperation.CancelJob, 3);
        request.AddGroup(IppGroupTag.Operation)
            .Add("attributes-charset", IppValue.OfString(IppValueTag.Charset, "utf-8"))
            .Add("job-id", IppValue.OfInteger(258));

        byte[] expected = Concat(
            new byte[] { 0x02, 0x00, 0x00, 0x08, 0x00, 0x00, 0x00, 0x03, 0x01 },
            Value(0x47, "attributes-charset", Ascii("utf-8")),
            Value(0x21, "job-id", new byte[] { 0x00, 0x00, 0x01, 0x02 }),
            new byte[] { 0x03 });
        Assert.Equal(expected, IppCodec.Encode(request));
    }

    [Fact]
    public void RoundTripsEveryValueTypeAndNestedCollections() {
        IppMessage original = Sample();

        IppMessage decoded = IppCodec.Decode(IppCodec.Encode(original));

        Assert.Equal(Describe(original), Describe(decoded));
        IppAttributeGroup job = decoded.FirstGroup(IppGroupTag.Job)!;
        Assert.True(job["when"]!.First!.TryGetDateTime(out DateTimeOffset when));
        Assert.Equal(new DateTimeOffset(2026, 11, 12, 15, 4, 5, 600, TimeSpan.FromHours(-5)), when);
        IppValue mediaCol = job["media-col"]!.First!;
        IppValue size = mediaCol.Members!.Single(m => m.Name == "media-size").First!;
        Assert.Equal(21000, size.Members!.Single(m => m.Name == "x-dimension").First!.Integer);
        Assert.Equal(new[] { "a", "b" }, mediaCol.Members!.Single(m => m.Name == "media-source").Values.Select(v => v.Text));
    }

    [Fact]
    public void RejectsEveryTruncationWithAProtocolException() {
        byte[] bytes = IppCodec.Encode(Sample());
        for (int length = 0; length < bytes.Length; length++) {
            Assert.Throws<IppException>(() => IppCodec.Decode(bytes, length, null, out _));
        }
    }

    [Theory]
    [InlineData(new byte[] { 0x02, 0x00, 0x00, 0x00, 0, 0, 0, 1, 0x00, 0x03 })] // reserved delimiter
    [InlineData(new byte[] { 0x02, 0x00, 0x00, 0x00, 0, 0, 0, 1, 0x21, 0, 1, (byte)'a', 0, 4, 0, 0, 0, 1, 0x03 })] // value before any group
    [InlineData(new byte[] { 0x02, 0x00, 0x00, 0x00, 0, 0, 0, 1, 0x01, 0x21, 0, 0, 0, 4, 0, 0, 0, 1, 0x03 })] // additional value without attribute
    [InlineData(new byte[] { 0x02, 0x00, 0x00, 0x00, 0, 0, 0, 1, 0x01, 0x21, 0, 1, (byte)'a', 0, 3, 0, 0, 1, 0x03 })] // short integer
    [InlineData(new byte[] { 0x02, 0x00, 0x00, 0x00, 0, 0, 0, 1, 0x01, 0x22, 0, 1, (byte)'a', 0, 1, 2, 0x03 })] // boolean 2
    [InlineData(new byte[] { 0x02, 0x00, 0x00, 0x00, 0, 0, 0, 1, 0x01, 0x44, 0, 1, (byte)'a', 0x80, 0x00, 0x03 })] // negative length
    [InlineData(new byte[] { 0x02, 0x00, 0x00, 0x00, 0, 0, 0, 1, 0x01, 0x37, 0, 1, (byte)'a', 0, 0, 0x03 })] // end-collection outside collection
    [InlineData(new byte[] { 0x02, 0x00, 0x00, 0x00, 0, 0, 0, 1, 0x01, 0x34, 0, 1, (byte)'a', 0, 0, 0x21, 0, 0, 0, 4, 0, 0, 0, 1, 0x37, 0, 0, 0, 0, 0x03 })] // member value without name
    [InlineData(new byte[] { 0x02, 0x00, 0x00, 0x00, 0, 0, 0, 1, 0x01, 0x34, 0, 1, (byte)'a', 0, 0, 0x03 })] // unterminated collection
    [InlineData(new byte[] { 0x02, 0x00, 0x00, 0x00, 0, 0, 0, 1, 0x01, 0x35, 0, 1, (byte)'a', 0, 5, 0, 9, 0, 0, 0, 0x03 })] // language overruns
    [InlineData(new byte[] { 0x02, 0x00, 0x00, 0x00, 0, 0, 0, 1, 0x01, 0x7F, 0, 1, (byte)'a', 0, 2, 0, 0, 0x03 })] // extension without tag
    [InlineData(new byte[] { 0x00, 0x00, 0x00, 0x00, 0, 0, 0, 1, 0x03 })] // version 0
    public void RejectsStructurallyInvalidMessages(byte[] bytes) {
        Assert.Throws<IppException>(() => IppCodec.Decode(bytes));
    }

    [Fact]
    public void RejectsCollectionsNestedBeyondTheLimit() {
        var body = new List<byte> { 0x02, 0x00, 0x00, 0x00, 0, 0, 0, 1, 0x01 };
        body.AddRange(Value(0x34, "c", Array.Empty<byte>()));
        for (int i = 0; i < 9; i++) {
            body.AddRange(Value(0x4A, "", Ascii("m")));
            body.AddRange(Value(0x34, "", Array.Empty<byte>()));
        }
        for (int i = 0; i < 10; i++) body.AddRange(Value(0x37, "", Array.Empty<byte>()));
        body.Add(0x03);
        byte[] bytes = body.ToArray();

        Assert.Throws<IppException>(() => IppCodec.Decode(bytes));
        IppMessage relaxed = IppCodec.Decode(bytes, new IppDecodeLimits { MaxCollectionDepth = 10 });
        Assert.Equal("c", relaxed.Groups[0].Attributes.Single().Name);
    }

    [Fact]
    public void EnforcesSizeAttributeAndValueCountLimits() {
        var message = new IppMessage(2, 0, 0, 1);
        IppAttributeGroup group = message.AddGroup(IppGroupTag.Printer);
        for (int i = 0; i < 20; i++) group.Add("a" + i, IppValue.OfInteger(i), IppValue.OfInteger(i));
        byte[] bytes = IppCodec.Encode(message);

        Assert.Throws<IppException>(() => IppCodec.Decode(bytes, new IppDecodeLimits { MaxMessageBytes = bytes.Length - 1 }));
        Assert.Throws<IppException>(() => IppCodec.Decode(bytes, new IppDecodeLimits { MaxAttributes = 19 }));
        Assert.Throws<IppException>(() => IppCodec.Decode(bytes, new IppDecodeLimits { MaxValues = 39 }));
        Assert.Throws<IppException>(() => IppCodec.Decode(bytes, new IppDecodeLimits { MaxGroups = 0 }));
        Assert.Equal(20, IppCodec.Decode(bytes, new IppDecodeLimits { MaxAttributes = 20, MaxValues = 40, MaxGroups = 1 }).Groups[0].Attributes.Count);
    }

    [Fact]
    public void KeepsUnknownTagsAndOutOfBandValuesWithoutFailing() {
        byte[] bytes = Concat(
            new byte[] { 0x02, 0x00, 0x00, 0x00, 0, 0, 0, 1, 0x0F },
            Value(0x13, "no-value", Array.Empty<byte>()),
            Value(0x12, "unknown", Array.Empty<byte>()),
            Value(0x38, "future-octets", new byte[] { 1, 2, 3 }),
            Value(0x7F, "extension", new byte[] { 0, 0, 1, 0, 9 }),
            new byte[] { 0x03 });

        IppMessage message = IppCodec.Decode(bytes);

        IppAttributeGroup group = Assert.Single(message.Groups);
        Assert.Equal((IppGroupTag)0x0F, group.Tag);
        Assert.True(group["no-value"]!.First!.IsOutOfBand);
        Assert.Equal(new byte[] { 1, 2, 3 }, group["future-octets"]!.First!.Octets);
        Assert.Equal(5, group["extension"]!.First!.Octets!.Length);
        Assert.Equal(bytes, IppCodec.Encode(message));
    }

    [Fact]
    public void InvalidDateTimeBytesAreReportedButNotThrown() {
        byte[] bytes = Concat(new byte[] { 0x02, 0x00, 0x00, 0x00, 0, 0, 0, 1, 0x04 },
            Value(0x31, "printer-current-time", new byte[] { 0x07, 0xEA, 13, 40, 25, 61, 61, 10, (byte)'*', 15, 60 }),
            new byte[] { 0x03 });
        IppValue value = IppCodec.Decode(bytes).Groups[0].Attributes[0].First!;
        Assert.False(value.TryGetDateTime(out _));
    }

    [Fact]
    public void BoundedFuzzLoopRaisesOnlyProtocolExceptions() {
        byte[] seed = IppCodec.Encode(Sample());
        var random = new Random(8010);
        var limits = new IppDecodeLimits { MaxMessageBytes = 64 * 1024, MaxAttributes = 512, MaxValues = 2048, MaxCollectionDepth = 4 };
        int accepted = 0;
        for (int iteration = 0; iteration < 20000; iteration++) {
            byte[] input = Mutate(seed, random);
            try {
                IppMessage message = IppCodec.Decode(input, input.Length, limits, out int offset);
                Assert.InRange(offset, 8, input.Length);
                // Anything the decoder accepts must re-encode and decode to the same structure.
                Assert.Equal(Describe(message), Describe(IppCodec.Decode(IppCodec.Encode(message), limits)));
                accepted++;
            } catch (IppException) {
            }
        }
        Assert.InRange(accepted, 1, 19999);
    }

    private static byte[] Mutate(byte[] seed, Random random) {
        byte[] data = (byte[])seed.Clone();
        int edits = random.Next(1, 6);
        for (int i = 0; i < edits; i++) {
            switch (random.Next(5)) {
                case 0: data[random.Next(data.Length)] = (byte)random.Next(256); break;
                case 1: data[random.Next(data.Length)] ^= (byte)(1 << random.Next(8)); break;
                case 2: Array.Resize(ref data, random.Next(data.Length + 1)); if (data.Length == 0) return data; break;
                case 3: {
                    int at = random.Next(data.Length);
                    data = data.Take(at).Concat(Enumerable.Repeat((byte)random.Next(256), random.Next(1, 64))).Concat(data.Skip(at)).ToArray();
                    break;
                }
                default: {
                    // Duplicate a slice so that groups, collections and values repeat or nest unexpectedly.
                    int from = random.Next(data.Length);
                    int count = random.Next(1, Math.Min(256, data.Length - from) + 1);
                    int at = random.Next(data.Length);
                    data = data.Take(at).Concat(data.Skip(from).Take(count)).Concat(data.Skip(at)).ToArray();
                    break;
                }
            }
        }
        return data;
    }

    internal static IppMessage Sample() {
        var message = new IppMessage(2, 0, (short)IppOperation.PrintJob, 42);
        message.AddGroup(IppGroupTag.Operation)
            .Add("attributes-charset", IppValue.OfString(IppValueTag.Charset, "utf-8"))
            .Add("attributes-natural-language", IppValue.OfString(IppValueTag.NaturalLanguage, "en"))
            .Add("printer-uri", IppValue.OfString(IppValueTag.Uri, "ipp://localhost/ipp/print"))
            .Add("job-name", IppValue.OfStringWithLanguage(IppValueTag.NameWithLanguage, "Raport kwartalny", "pl"))
            .Add("document-format", IppValue.OfString(IppValueTag.MimeMediaType, "application/pdf"));
        message.AddGroup(IppGroupTag.Job)
            .Add("copies", IppValue.OfInteger(2))
            .Add("orientation-requested", IppValue.OfEnum(4))
            .Add("ipp-attribute-fidelity", IppValue.OfBoolean(true))
            .Add("page-ranges", IppValue.OfRange(1, 3), IppValue.OfRange(5, 5))
            .Add("printer-resolution", IppValue.OfResolution(600, 1200, 3))
            .Add("when", IppValue.OfDateTime(new DateTimeOffset(2026, 11, 12, 15, 4, 5, 600, TimeSpan.FromHours(-5))))
            .Add("job-password", IppValue.OfOctets(new byte[] { 0, 1, 255 }))
            .Add("job-message", IppValue.OfStringWithLanguage(IppValueTag.TextWithLanguage, "Zażółć gęślą jaźń", "pl-PL"), IppValue.OfString(IppValueTag.TextWithoutLanguage, "plain"))
            .Add("sides", IppValue.OfString(IppValueTag.Keyword, "two-sided-long-edge"))
            .Add("job-sheets", IppValue.OutOfBand(IppValueTag.NoValue))
            .Add("media-col", IppValue.OfCollection(new[] {
                new IppAttribute("media-size", IppValue.OfCollection(new[] {
                    new IppAttribute("x-dimension", IppValue.OfInteger(21000)),
                    new IppAttribute("y-dimension", IppValue.OfInteger(29700))
                })),
                new IppAttribute("media-source", IppValue.OfString(IppValueTag.Keyword, "a"), IppValue.OfString(IppValueTag.Keyword, "b"))
            }), IppValue.OfCollection(Array.Empty<IppAttribute>()));
        return message;
    }

    internal static string Describe(IppMessage message) =>
        $"{message.VersionMajor}.{message.VersionMinor} {message.Code} {message.RequestId}|" +
        string.Join("|", message.Groups.Select(group => ((byte)group.Tag).ToString() + ":" + string.Join(";", group.Attributes.Select(Describe))));

    private static string Describe(IppAttribute attribute) => attribute.Name + "=" + string.Join(",", attribute.Values.Select(Describe));

    private static string Describe(IppValue value) => ((byte)value.Tag).ToString("X2") + "(" + value.Tag switch {
        IppValueTag.BegCollection => string.Join(";", value.Members!.Select(Describe)),
        _ => $"{value.Integer},{value.Second},{value.Units},{value.Text},{value.Language},{(value.Octets is null ? "" : Convert.ToBase64String(value.Octets))}"
    } + ")";

    internal static byte[] Value(byte tag, string name, byte[] value) {
        byte[] nameBytes = Ascii(name);
        return Concat(new[] { tag, (byte)(nameBytes.Length >> 8), (byte)nameBytes.Length }, nameBytes,
            new[] { (byte)(value.Length >> 8), (byte)value.Length }, value);
    }

    internal static byte[] Ascii(string value) => System.Text.Encoding.ASCII.GetBytes(value);

    internal static byte[] Concat(params byte[][] parts) => parts.SelectMany(part => part).ToArray();
}
