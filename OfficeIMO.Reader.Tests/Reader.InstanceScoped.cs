using OfficeIMO.Reader;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class ReaderInstanceScopedTests {
    [Fact]
    public async Task OfficeDocumentReader_IsolatesHandlersAcrossConcurrentInstances() {
        const string extension = ".instanceix";
        string file = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);
        File.WriteAllText(file, "input");

        try {
            OfficeDocumentReader first = BuildReader("officeimo.tests.instance.first", extension, "first-output");
            OfficeDocumentReader second = BuildReader("officeimo.tests.instance.second", extension, "second-output");

            Task<ReaderChunk> firstRead = Task.Run(() => Assert.Single(first.Read(file)));
            Task<ReaderChunk> secondRead = Task.Run(() => Assert.Single(second.Read(file)));
            ReaderChunk[] chunks = await Task.WhenAll(firstRead, secondRead);

            Assert.Equal("first-output", chunks[0].Text);
            Assert.Equal("second-output", chunks[1].Text);
            Assert.Equal("officeimo.tests.instance.first", Assert.Single(first.GetCapabilities(false, true)).Id);
            Assert.Equal("officeimo.tests.instance.second", Assert.Single(second.GetCapabilities(false, true)).Id);
        } finally {
            File.Delete(file);
        }
    }

    [Fact]
    public void OfficeDocumentReader_BuildCapturesImmutableHandlerSnapshot() {
        const string extension = ".snapshotix";
        string file = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);
        File.WriteAllText(file, "input");

        try {
            var builder = new OfficeDocumentReaderBuilder();
            builder.AddHandler(CreateRegistration("officeimo.tests.snapshot.first", extension, "before-build"));
            OfficeDocumentReader first = builder.Build();

            builder.AddHandler(
                CreateRegistration("officeimo.tests.snapshot.second", extension, "after-build"),
                replaceExisting: true);
            OfficeDocumentReader second = builder.Build();

            Assert.Equal("before-build", Assert.Single(first.Read(file)).Text);
            Assert.Equal("after-build", Assert.Single(second.Read(file)).Text);
            Assert.Equal("officeimo.tests.snapshot.first", Assert.Single(first.GetCapabilities(false, true)).Id);
            Assert.Equal("officeimo.tests.snapshot.second", Assert.Single(second.GetCapabilities(false, true)).Id);
        } finally {
            File.Delete(file);
        }
    }

    [Fact]
    public void OfficeDocumentReader_PreservesRichResultForInstanceHandler() {
        const string extension = ".instancerichix";
        var builder = new OfficeDocumentReaderBuilder();
        builder.AddHandler(new ReaderHandlerRegistration {
            Id = "officeimo.tests.instance.rich",
            Kind = ReaderInputKind.Text,
            Extensions = new[] { extension },
            ReadDocumentStream = (stream, sourceName, options, cancellationToken) => new OfficeDocumentReadResult {
                Kind = ReaderInputKind.Text,
                Source = new OfficeDocumentSource { Path = sourceName },
                CapabilitiesUsed = new[] { "officeimo.tests.instance.rich" },
                Chunks = new[] {
                    new ReaderChunk {
                        Id = "instance-rich-0001",
                        Kind = ReaderInputKind.Text,
                        Text = "rich-instance-output"
                    }
                },
                Links = new[] {
                    new OfficeDocumentLink {
                        Id = "instance-link-0001",
                        Kind = "external",
                        Uri = "https://example.test/instance"
                    }
                }
            }
        });
        OfficeDocumentReader reader = builder.Build();

        using var stream = new MemoryStream(new byte[] { 1, 2, 3 });
        OfficeDocumentReadResult result = reader.ReadDocument(stream, "sample" + extension);

        Assert.Equal("rich-instance-output", Assert.Single(result.Chunks).Text);
        Assert.Equal("https://example.test/instance", Assert.Single(result.Links).Uri);
        Assert.True(Assert.Single(reader.GetCapabilities(false, true)).SupportsDocumentStream);
    }

    [Fact]
    public void OfficeDocumentReader_KeepsScopeCorrectForInterleavedEnumerators() {
        const string extension = ".interleavedix";
        string file = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);
        File.WriteAllText(file, "input");

        try {
            OfficeDocumentReader first = BuildDetectingReader("officeimo.tests.interleaved.first", extension, ReaderInputKind.Csv);
            OfficeDocumentReader second = BuildDetectingReader("officeimo.tests.interleaved.second", extension, ReaderInputKind.Json);

            using IEnumerator<ReaderChunk> firstChunks = first.Read(file).GetEnumerator();
            using IEnumerator<ReaderChunk> secondChunks = second.Read(file).GetEnumerator();

            Assert.True(firstChunks.MoveNext());
            Assert.Equal(nameof(ReaderInputKind.Csv), firstChunks.Current.Text);
            Assert.True(secondChunks.MoveNext());
            Assert.Equal(nameof(ReaderInputKind.Json), secondChunks.Current.Text);
            Assert.True(firstChunks.MoveNext());
            Assert.Equal(nameof(ReaderInputKind.Csv), firstChunks.Current.Text);
            Assert.True(secondChunks.MoveNext());
            Assert.Equal(nameof(ReaderInputKind.Json), secondChunks.Current.Text);
        } finally {
            File.Delete(file);
        }
    }

    private static OfficeDocumentReader BuildReader(string id, string extension, string output) {
        return new OfficeDocumentReaderBuilder()
            .AddHandler(CreateRegistration(id, extension, output))
            .Build();
    }

    private static ReaderHandlerRegistration CreateRegistration(string id, string extension, string output) {
        return new ReaderHandlerRegistration {
            Id = id,
            Kind = ReaderInputKind.Text,
            Extensions = new[] { extension },
            ReadPath = (path, options, cancellationToken) => new[] {
                new ReaderChunk {
                    Id = id + ":0001",
                    Kind = ReaderInputKind.Text,
                    Location = new ReaderLocation { Path = path },
                    Text = output
                }
            }
        };
    }

    private static OfficeDocumentReader BuildDetectingReader(string id, string extension, ReaderInputKind kind) {
        return new OfficeDocumentReaderBuilder()
            .AddHandler(new ReaderHandlerRegistration {
                Id = id,
                Kind = kind,
                Extensions = new[] { extension },
                ReadPath = (path, options, cancellationToken) => DetectKindForEachChunk(extension)
            })
            .Build();
    }

    private static IEnumerable<ReaderChunk> DetectKindForEachChunk(string extension) {
        for (int index = 0; index < 2; index++) {
            ReaderInputKind detected = DocumentReader.DetectKind("nested" + extension);
            yield return new ReaderChunk {
                Id = "nested:" + index,
                Kind = detected,
                Text = detected.ToString()
            };
        }
    }
}

public sealed partial class ReaderRegistryTests {
    [Fact]
    public void OfficeDocumentReader_IsNotAffectedByStaticRegistration() {
        const string staticHandlerId = "officeimo.tests.instance.static";
        const string instanceHandlerId = "officeimo.tests.instance.isolated";
        const string extension = ".staticisolationix";
        string file = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);

        try {
            DocumentReader.UnregisterHandler(staticHandlerId);
            DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
                Id = staticHandlerId,
                Kind = ReaderInputKind.Text,
                Extensions = new[] { extension },
                ReadPath = (path, options, cancellationToken) => new[] {
                    new ReaderChunk { Id = "static", Kind = ReaderInputKind.Text, Text = "static-output" }
                }
            });

            OfficeDocumentReader reader = new OfficeDocumentReaderBuilder()
                .AddHandler(new ReaderHandlerRegistration {
                    Id = instanceHandlerId,
                    Kind = ReaderInputKind.Text,
                    Extensions = new[] { extension },
                    ReadPath = (path, options, cancellationToken) => new[] {
                        new ReaderChunk { Id = "instance", Kind = ReaderInputKind.Text, Text = "instance-output" }
                    }
                })
                .Build();
            File.WriteAllText(file, "input");

            Assert.Equal("static-output", Assert.Single(DocumentReader.Read(file)).Text);
            Assert.Equal("instance-output", Assert.Single(reader.Read(file)).Text);
            Assert.Equal(instanceHandlerId, Assert.Single(reader.GetCapabilities(false, true)).Id);
        } finally {
            DocumentReader.UnregisterHandler(staticHandlerId);
            File.Delete(file);
        }
    }
}
