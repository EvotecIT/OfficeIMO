using OfficeIMO.Reader;
using OfficeIMO.Reader.Csv;
using OfficeIMO.Reader.Epub;
using OfficeIMO.Reader.Html;
using OfficeIMO.Reader.Json;
using OfficeIMO.Reader.Text;
using OfficeIMO.Reader.Xml;
using OfficeIMO.Reader.Zip;
using System.IO.Compression;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests;

[CollectionDefinition("ReaderRegistryNonParallel", DisableParallelization = true)]
public sealed class ReaderRegistryNonParallelCollection;

[Collection("ReaderRegistryNonParallel")]
public sealed class ReaderRegistryTests {
    [Fact]
    public void DocumentReader_GetCapabilities_IncludesBuiltInHandlers() {
        var capabilities = DocumentReader.GetCapabilities();

        Assert.NotEmpty(capabilities);
        Assert.All(capabilities, capability => {
            Assert.Equal(ReaderCapabilitySchema.Id, capability.SchemaId);
            Assert.Equal(ReaderCapabilitySchema.Version, capability.SchemaVersion);
        });
        Assert.Contains(capabilities, c => c.IsBuiltIn && c.Id == "officeimo.reader.word");
        Assert.Contains(capabilities, c => c.IsBuiltIn && c.Id == "officeimo.reader.excel");
        Assert.Contains(capabilities, c => c.IsBuiltIn && c.Id == "officeimo.reader.powerpoint");
        Assert.Contains(capabilities, c => c.IsBuiltIn && c.Id == "officeimo.reader.markdown");
        Assert.Contains(capabilities, c => c.IsBuiltIn && c.Id == "officeimo.reader.pdf");
        Assert.Contains(capabilities, c => c.IsBuiltIn && c.Id == "officeimo.reader.text");
    }

    [Fact]
    public void DocumentReader_GetCapabilityManifest_ReflectsCapabilities() {
        var manifest = DocumentReader.GetCapabilityManifest();
        var capabilities = DocumentReader.GetCapabilities();

        Assert.NotNull(manifest);
        Assert.Equal(ReaderCapabilitySchema.Id, manifest.SchemaId);
        Assert.Equal(ReaderCapabilitySchema.Version, manifest.SchemaVersion);
        Assert.Equal(capabilities.Count, manifest.Handlers.Count);
        Assert.Contains(manifest.Handlers, c => c.Id == "officeimo.reader.word");
    }

    [Fact]
    public void DocumentReader_GetCapabilityManifestJson_IsDeterministicAndValidJson() {
        var jsonA = DocumentReader.GetCapabilityManifestJson(indented: false);
        var jsonB = DocumentReader.GetCapabilityManifestJson(indented: false);
        Assert.Equal(jsonA, jsonB);

        using var stream = new MemoryStream(Encoding.UTF8.GetBytes(jsonA));
        var chunks = DocumentReaderJsonExtensions.ReadJson(
            stream,
            sourceName: "capability-manifest.json",
            jsonOptions: new JsonReadOptions {
                ChunkRows = 128,
                IncludeMarkdown = false
            }).ToList();

        Assert.NotEmpty(chunks);
        Assert.DoesNotContain(chunks, c =>
            c.Warnings?.Any(w => w.Contains("JSON parse error", StringComparison.OrdinalIgnoreCase)) ?? false);
        Assert.Contains(chunks, c =>
            (c.Text?.Contains("$.schemaId", StringComparison.Ordinal) ?? false) &&
            (c.Text?.Contains("officeimo.reader.capability", StringComparison.Ordinal) ?? false));
        Assert.Contains(chunks, c =>
            (c.Text?.Contains("$.handlers[0].id", StringComparison.Ordinal) ?? false));
    }

    [Fact]
    public void DocumentReader_RegisterHandler_UsesCustomPathReader() {
        const string handlerId = "officeimo.tests.custom.demo";
        const string extension = ".demoix";

        var file = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + extension);
        try {
            DocumentReader.UnregisterHandler(handlerId);

            DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
                Id = handlerId,
                DisplayName = "Test custom handler",
                Kind = ReaderInputKind.Text,
                Extensions = new[] { extension },
                DefaultMaxInputBytes = 4096,
                WarningBehavior = ReaderWarningBehavior.WarningChunksOnly,
                DeterministicOutput = false,
                ReadPath = (path, options, ct) => new[] {
                    new ReaderChunk {
                        Id = "custom-0001",
                        Kind = ReaderInputKind.Text,
                        Location = new ReaderLocation {
                            Path = path,
                            BlockIndex = 0
                        },
                        Text = "custom-handler-output"
                    }
                }
            });

            File.WriteAllText(file, "input");

            var kind = DocumentReader.DetectKind(file);
            Assert.Equal(ReaderInputKind.Text, kind);

            var chunks = DocumentReader.Read(file).ToList();
            Assert.Single(chunks);
            Assert.Equal("custom-handler-output", chunks[0].Text);

            var customCapabilities = DocumentReader.GetCapabilities(includeBuiltIn: false, includeCustom: true);
            Assert.Contains(customCapabilities, c =>
                c.Id == handlerId &&
                c.Extensions.Contains(extension, StringComparer.OrdinalIgnoreCase) &&
                c.DefaultMaxInputBytes == 4096 &&
                c.WarningBehavior == ReaderWarningBehavior.WarningChunksOnly &&
                c.DeterministicOutput == false &&
                c.SchemaId == ReaderCapabilitySchema.Id &&
                c.SchemaVersion == ReaderCapabilitySchema.Version);
        } finally {
            DocumentReader.UnregisterHandler(handlerId);
            if (File.Exists(file)) File.Delete(file);
        }
    }

    [Fact]
    public void DocumentReader_RegisterHandler_RejectsInvalidAdvertisedDefaultMaxInputBytes() {
        const string handlerId = "officeimo.tests.custom.invalidmax";
        DocumentReader.UnregisterHandler(handlerId);

        try {
            Assert.Throws<ArgumentException>(() => DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
                Id = handlerId,
                Extensions = new[] { ".invalidmax" },
                Kind = ReaderInputKind.Unknown,
                DefaultMaxInputBytes = 0,
                ReadPath = (path, options, ct) => Array.Empty<ReaderChunk>()
            }));
        } finally {
            DocumentReader.UnregisterHandler(handlerId);
        }
    }

    [Fact]
    public void DocumentReader_RegisterHandler_WithoutReplaceExisting_RejectsBuiltInCollision() {
        const string handlerId = "officeimo.tests.custom.markdown";

        DocumentReader.UnregisterHandler(handlerId);
        try {
            Assert.Throws<InvalidOperationException>(() => DocumentReader.RegisterHandler(new ReaderHandlerRegistration {
                Id = handlerId,
                Extensions = new[] { ".md" },
                Kind = ReaderInputKind.Markdown,
                ReadPath = (path, options, ct) => Array.Empty<ReaderChunk>()
            }));
        } finally {
            DocumentReader.UnregisterHandler(handlerId);
        }
    }

    [Fact]
    public void DocumentReader_ModularRegistrationHelpers_RegisterAndUnregister() {
        try {
            DocumentReaderCsvRegistrationExtensions.RegisterCsvHandler(replaceExisting: true);
            DocumentReaderEpubRegistrationExtensions.RegisterEpubHandler(replaceExisting: true);
            DocumentReaderZipRegistrationExtensions.RegisterZipHandler(replaceExisting: true);
            DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler(replaceExisting: true);
            DocumentReaderJsonRegistrationExtensions.RegisterJsonHandler(replaceExisting: true);
            DocumentReaderXmlRegistrationExtensions.RegisterXmlHandler(replaceExisting: true);

            var capabilities = DocumentReader.GetCapabilities();
            var csvCapability = Assert.Single(capabilities, c => c.Id == DocumentReaderCsvRegistrationExtensions.HandlerId);
            var epubCapability = Assert.Single(capabilities, c => c.Id == DocumentReaderEpubRegistrationExtensions.HandlerId);
            var zipCapability = Assert.Single(capabilities, c => c.Id == DocumentReaderZipRegistrationExtensions.HandlerId);
            var htmlCapability = Assert.Single(capabilities, c => c.Id == DocumentReaderHtmlRegistrationExtensions.HandlerId);
            var jsonCapability = Assert.Single(capabilities, c => c.Id == DocumentReaderJsonRegistrationExtensions.HandlerId);
            var xmlCapability = Assert.Single(capabilities, c => c.Id == DocumentReaderXmlRegistrationExtensions.HandlerId);

            Assert.True(csvCapability.SupportsPath);
            Assert.True(csvCapability.SupportsStream);
            Assert.True(epubCapability.SupportsPath);
            Assert.True(epubCapability.SupportsStream);
            Assert.True(zipCapability.SupportsPath);
            Assert.True(zipCapability.SupportsStream);
            Assert.True(htmlCapability.SupportsPath);
            Assert.True(htmlCapability.SupportsStream);
            Assert.True(jsonCapability.SupportsPath);
            Assert.True(jsonCapability.SupportsStream);
            Assert.True(xmlCapability.SupportsPath);
            Assert.True(xmlCapability.SupportsStream);
            Assert.Equal(ReaderCapabilitySchema.Id, csvCapability.SchemaId);
            Assert.Equal(ReaderCapabilitySchema.Version, csvCapability.SchemaVersion);
            Assert.Equal(ReaderWarningBehavior.Mixed, csvCapability.WarningBehavior);
            Assert.True(csvCapability.DeterministicOutput);
            Assert.Equal(ReaderCapabilitySchema.Id, epubCapability.SchemaId);
            Assert.Equal(ReaderCapabilitySchema.Version, epubCapability.SchemaVersion);
            Assert.Equal(ReaderWarningBehavior.Mixed, epubCapability.WarningBehavior);
            Assert.True(epubCapability.DeterministicOutput);
            Assert.Equal(ReaderCapabilitySchema.Id, zipCapability.SchemaId);
            Assert.Equal(ReaderCapabilitySchema.Version, zipCapability.SchemaVersion);
            Assert.Equal(ReaderWarningBehavior.Mixed, zipCapability.WarningBehavior);
            Assert.True(zipCapability.DeterministicOutput);
            Assert.Equal(ReaderCapabilitySchema.Id, htmlCapability.SchemaId);
            Assert.Equal(ReaderCapabilitySchema.Version, htmlCapability.SchemaVersion);
            Assert.Equal(ReaderWarningBehavior.Mixed, htmlCapability.WarningBehavior);
            Assert.True(htmlCapability.DeterministicOutput);
            Assert.Equal(ReaderCapabilitySchema.Id, jsonCapability.SchemaId);
            Assert.Equal(ReaderCapabilitySchema.Version, jsonCapability.SchemaVersion);
            Assert.Equal(ReaderWarningBehavior.Mixed, jsonCapability.WarningBehavior);
            Assert.True(jsonCapability.DeterministicOutput);
            Assert.Equal(ReaderCapabilitySchema.Id, xmlCapability.SchemaId);
            Assert.Equal(ReaderCapabilitySchema.Version, xmlCapability.SchemaVersion);
            Assert.Equal(ReaderWarningBehavior.Mixed, xmlCapability.WarningBehavior);
            Assert.True(xmlCapability.DeterministicOutput);

            Assert.Equal(ReaderInputKind.Epub, DocumentReader.DetectKind("book.epub"));
            Assert.Equal(ReaderInputKind.Zip, DocumentReader.DetectKind("archive.zip"));
            Assert.Equal(ReaderInputKind.Html, DocumentReader.DetectKind("index.html"));
            Assert.Equal(ReaderInputKind.Json, DocumentReader.DetectKind("data.json"));
        } finally {
            DocumentReaderCsvRegistrationExtensions.UnregisterCsvHandler();
            DocumentReaderEpubRegistrationExtensions.UnregisterEpubHandler();
            DocumentReaderZipRegistrationExtensions.UnregisterZipHandler();
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
            DocumentReaderJsonRegistrationExtensions.UnregisterJsonHandler();
            DocumentReaderXmlRegistrationExtensions.UnregisterXmlHandler();
        }
    }

    [Fact]
    public void DocumentReader_ReadFolder_DefaultExtensions_IncludeRegisteredCustomHandlers() {
        var folder = Path.Combine(Path.GetTempPath(), "officeimo-reader-folder-" + Guid.NewGuid().ToString("N"));
        var htmlPath = Path.Combine(folder, "index.html");

        Directory.CreateDirectory(folder);
        File.WriteAllText(htmlPath, "<html><body><h1>Folder HTML</h1><p>Body</p></body></html>");

        try {
            DocumentReaderHtmlRegistrationExtensions.RegisterHtmlHandler(replaceExisting: true);

            var chunks = DocumentReader.ReadFolder(
                folderPath: folder,
                folderOptions: new ReaderFolderOptions {
                    Recurse = false,
                    MaxFiles = 10
                },
                options: new ReaderOptions {
                    MaxChars = 8_000
                }).ToList();

            Assert.NotEmpty(chunks);
            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Html &&
                string.Equals(c.Location.Path, htmlPath, StringComparison.OrdinalIgnoreCase) &&
                ((c.Markdown ?? c.Text).Contains("Folder HTML", StringComparison.Ordinal) ||
                 (c.Markdown ?? c.Text).Contains("Body", StringComparison.Ordinal)));
        } finally {
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
            if (Directory.Exists(folder)) {
                Directory.Delete(folder, recursive: true);
            }
        }
    }

    [Fact]
    public void DocumentReader_DiscoverHandlerRegistrars_FindsModularRegistrars() {
        var registrars = DocumentReader.DiscoverHandlerRegistrars(
            typeof(DocumentReaderCsvRegistrationExtensions).Assembly,
            typeof(DocumentReaderEpubRegistrationExtensions).Assembly,
            typeof(DocumentReaderZipRegistrationExtensions).Assembly,
            typeof(DocumentReaderHtmlRegistrationExtensions).Assembly,
            typeof(DocumentReaderJsonRegistrationExtensions).Assembly,
            typeof(DocumentReaderXmlRegistrationExtensions).Assembly).ToList();

        Assert.NotEmpty(registrars);
        Assert.Contains(registrars, r => r.HandlerId == DocumentReaderCsvRegistrationExtensions.HandlerId);
        Assert.Contains(registrars, r => r.HandlerId == DocumentReaderEpubRegistrationExtensions.HandlerId);
        Assert.Contains(registrars, r => r.HandlerId == DocumentReaderZipRegistrationExtensions.HandlerId);
        Assert.Contains(registrars, r => r.HandlerId == DocumentReaderHtmlRegistrationExtensions.HandlerId);
        Assert.Contains(registrars, r => r.HandlerId == DocumentReaderJsonRegistrationExtensions.HandlerId);
        Assert.Contains(registrars, r => r.HandlerId == DocumentReaderXmlRegistrationExtensions.HandlerId);

        var ordered = registrars
            .OrderBy(r => r.HandlerId, StringComparer.Ordinal)
            .ThenBy(r => r.AssemblyName, StringComparer.Ordinal)
            .ThenBy(r => r.TypeName, StringComparer.Ordinal)
            .ThenBy(r => r.MethodName, StringComparer.Ordinal)
            .ToList();
        Assert.Equal(
            ordered.Select(static r => r.HandlerId).ToArray(),
            registrars.Select(static r => r.HandlerId).ToArray());
    }

    [Fact]
    public void DocumentReader_RegisterHandlersFromAssemblies_RegistersModularHandlers() {
        try {
            DocumentReaderCsvRegistrationExtensions.UnregisterCsvHandler();
            DocumentReaderEpubRegistrationExtensions.UnregisterEpubHandler();
            DocumentReaderZipRegistrationExtensions.UnregisterZipHandler();
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
            DocumentReaderJsonRegistrationExtensions.UnregisterJsonHandler();
            DocumentReaderXmlRegistrationExtensions.UnregisterXmlHandler();

            var registered = DocumentReader.RegisterHandlersFromAssemblies(
                replaceExisting: true,
                typeof(DocumentReaderCsvRegistrationExtensions).Assembly,
                typeof(DocumentReaderEpubRegistrationExtensions).Assembly,
                typeof(DocumentReaderZipRegistrationExtensions).Assembly,
                typeof(DocumentReaderHtmlRegistrationExtensions).Assembly,
                typeof(DocumentReaderJsonRegistrationExtensions).Assembly,
                typeof(DocumentReaderXmlRegistrationExtensions).Assembly).ToList();

            Assert.NotEmpty(registered);
            Assert.Contains(registered, r => r.HandlerId == DocumentReaderCsvRegistrationExtensions.HandlerId);
            Assert.Contains(registered, r => r.HandlerId == DocumentReaderEpubRegistrationExtensions.HandlerId);
            Assert.Contains(registered, r => r.HandlerId == DocumentReaderZipRegistrationExtensions.HandlerId);
            Assert.Contains(registered, r => r.HandlerId == DocumentReaderHtmlRegistrationExtensions.HandlerId);
            Assert.Contains(registered, r => r.HandlerId == DocumentReaderJsonRegistrationExtensions.HandlerId);
            Assert.Contains(registered, r => r.HandlerId == DocumentReaderXmlRegistrationExtensions.HandlerId);

            var capabilities = DocumentReader.GetCapabilities();
            Assert.Contains(capabilities, c => c.Id == DocumentReaderCsvRegistrationExtensions.HandlerId);
            Assert.Contains(capabilities, c => c.Id == DocumentReaderEpubRegistrationExtensions.HandlerId);
            Assert.Contains(capabilities, c => c.Id == DocumentReaderZipRegistrationExtensions.HandlerId);
            Assert.Contains(capabilities, c => c.Id == DocumentReaderHtmlRegistrationExtensions.HandlerId);
            Assert.Contains(capabilities, c => c.Id == DocumentReaderJsonRegistrationExtensions.HandlerId);
            Assert.Contains(capabilities, c => c.Id == DocumentReaderXmlRegistrationExtensions.HandlerId);
        } finally {
            DocumentReaderCsvRegistrationExtensions.UnregisterCsvHandler();
            DocumentReaderEpubRegistrationExtensions.UnregisterEpubHandler();
            DocumentReaderZipRegistrationExtensions.UnregisterZipHandler();
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
            DocumentReaderJsonRegistrationExtensions.UnregisterJsonHandler();
            DocumentReaderXmlRegistrationExtensions.UnregisterXmlHandler();
        }
    }

    [Fact]
    public void DocumentReader_DiscoverHandlerRegistrarsFromLoadedAssemblies_FindsModularRegistrars() {
        EnsureModularReaderAssembliesLoaded();

        var registrars = DocumentReader.DiscoverHandlerRegistrarsFromLoadedAssemblies().ToList();

        Assert.NotEmpty(registrars);
        Assert.Contains(registrars, r => r.HandlerId == DocumentReaderCsvRegistrationExtensions.HandlerId);
        Assert.Contains(registrars, r => r.HandlerId == DocumentReaderEpubRegistrationExtensions.HandlerId);
        Assert.Contains(registrars, r => r.HandlerId == DocumentReaderZipRegistrationExtensions.HandlerId);
        Assert.Contains(registrars, r => r.HandlerId == DocumentReaderHtmlRegistrationExtensions.HandlerId);
        Assert.Contains(registrars, r => r.HandlerId == DocumentReaderJsonRegistrationExtensions.HandlerId);
        Assert.Contains(registrars, r => r.HandlerId == DocumentReaderXmlRegistrationExtensions.HandlerId);
        Assert.DoesNotContain(registrars, r => r.HandlerId == DocumentReaderTextRegistrationExtensions.HandlerId);
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("   ")]
    public void DocumentReader_DiscoverHandlerRegistrarsFromLoadedAssemblies_EmptyPrefix_Throws(string prefix) {
        Assert.Throws<ArgumentException>(() => DocumentReader.DiscoverHandlerRegistrarsFromLoadedAssemblies(prefix));
    }

    [Fact]
    public void DocumentReader_DiscoverHandlerRegistrarsFromLoadedAssemblies_NoMatches_ReturnsEmpty() {
        EnsureModularReaderAssembliesLoaded();

        var registrars = DocumentReader.DiscoverHandlerRegistrarsFromLoadedAssemblies("OfficeIMO.Reader.DoesNotExist.");
        Assert.Empty(registrars);
    }

    [Fact]
    public void DocumentReader_RegisterHandlersFromLoadedAssemblies_RegistersModularHandlers() {
        EnsureModularReaderAssembliesLoaded();

        try {
            DocumentReaderCsvRegistrationExtensions.UnregisterCsvHandler();
            DocumentReaderEpubRegistrationExtensions.UnregisterEpubHandler();
            DocumentReaderZipRegistrationExtensions.UnregisterZipHandler();
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
            DocumentReaderJsonRegistrationExtensions.UnregisterJsonHandler();
            DocumentReaderXmlRegistrationExtensions.UnregisterXmlHandler();

            var registered = DocumentReader.RegisterHandlersFromLoadedAssemblies(replaceExisting: true).ToList();

            Assert.NotEmpty(registered);
            Assert.Contains(registered, r => r.HandlerId == DocumentReaderCsvRegistrationExtensions.HandlerId);
            Assert.Contains(registered, r => r.HandlerId == DocumentReaderEpubRegistrationExtensions.HandlerId);
            Assert.Contains(registered, r => r.HandlerId == DocumentReaderZipRegistrationExtensions.HandlerId);
            Assert.Contains(registered, r => r.HandlerId == DocumentReaderHtmlRegistrationExtensions.HandlerId);
            Assert.Contains(registered, r => r.HandlerId == DocumentReaderJsonRegistrationExtensions.HandlerId);
            Assert.Contains(registered, r => r.HandlerId == DocumentReaderXmlRegistrationExtensions.HandlerId);

            var capabilities = DocumentReader.GetCapabilities();
            Assert.Contains(capabilities, c => c.Id == DocumentReaderCsvRegistrationExtensions.HandlerId);
            Assert.Contains(capabilities, c => c.Id == DocumentReaderEpubRegistrationExtensions.HandlerId);
            Assert.Contains(capabilities, c => c.Id == DocumentReaderZipRegistrationExtensions.HandlerId);
            Assert.Contains(capabilities, c => c.Id == DocumentReaderHtmlRegistrationExtensions.HandlerId);
            Assert.Contains(capabilities, c => c.Id == DocumentReaderJsonRegistrationExtensions.HandlerId);
            Assert.Contains(capabilities, c => c.Id == DocumentReaderXmlRegistrationExtensions.HandlerId);
        } finally {
            DocumentReaderCsvRegistrationExtensions.UnregisterCsvHandler();
            DocumentReaderEpubRegistrationExtensions.UnregisterEpubHandler();
            DocumentReaderZipRegistrationExtensions.UnregisterZipHandler();
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
            DocumentReaderJsonRegistrationExtensions.UnregisterJsonHandler();
            DocumentReaderXmlRegistrationExtensions.UnregisterXmlHandler();
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("   ")]
    public void DocumentReader_RegisterHandlersFromLoadedAssemblies_EmptyPrefix_Throws(string prefix) {
        Assert.Throws<ArgumentException>(() => DocumentReader.RegisterHandlersFromLoadedAssemblies(replaceExisting: true, assemblyNamePrefix: prefix));
    }

    [Fact]
    public void DocumentReader_BootstrapHostFromLoadedAssemblies_RegistersHandlersAndBuildsManifestPayload() {
        EnsureModularReaderAssembliesLoaded();

        try {
            DocumentReaderCsvRegistrationExtensions.UnregisterCsvHandler();
            DocumentReaderEpubRegistrationExtensions.UnregisterEpubHandler();
            DocumentReaderZipRegistrationExtensions.UnregisterZipHandler();
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
            DocumentReaderJsonRegistrationExtensions.UnregisterJsonHandler();
            DocumentReaderXmlRegistrationExtensions.UnregisterXmlHandler();

            var result = DocumentReader.BootstrapHostFromLoadedAssemblies(options: new ReaderHostBootstrapOptions {
                ReplaceExistingHandlers = true,
                IncludeBuiltInCapabilities = true,
                IncludeCustomCapabilities = true,
                IndentedManifestJson = false
            });

            Assert.NotNull(result);
            Assert.Equal("OfficeIMO.Reader.", result.AssemblyNamePrefix);
            Assert.True(result.ReplaceExistingHandlers);
            Assert.NotEmpty(result.RegisteredHandlers);
            Assert.Contains(result.RegisteredHandlers, r => r.HandlerId == DocumentReaderCsvRegistrationExtensions.HandlerId);
            Assert.Contains(result.RegisteredHandlers, r => r.HandlerId == DocumentReaderEpubRegistrationExtensions.HandlerId);
            Assert.Contains(result.RegisteredHandlers, r => r.HandlerId == DocumentReaderZipRegistrationExtensions.HandlerId);
            Assert.Contains(result.RegisteredHandlers, r => r.HandlerId == DocumentReaderHtmlRegistrationExtensions.HandlerId);
            Assert.Contains(result.RegisteredHandlers, r => r.HandlerId == DocumentReaderJsonRegistrationExtensions.HandlerId);
            Assert.Contains(result.RegisteredHandlers, r => r.HandlerId == DocumentReaderXmlRegistrationExtensions.HandlerId);

            Assert.Equal(ReaderCapabilitySchema.Id, result.Manifest.SchemaId);
            Assert.Equal(ReaderCapabilitySchema.Version, result.Manifest.SchemaVersion);
            Assert.Contains(result.Manifest.Handlers, c => c.Id == "officeimo.reader.word");
            Assert.Contains(result.Manifest.Handlers, c => c.Id == DocumentReaderCsvRegistrationExtensions.HandlerId);
            Assert.Contains(result.Manifest.Handlers, c => c.Id == DocumentReaderEpubRegistrationExtensions.HandlerId);
            Assert.Contains(result.Manifest.Handlers, c => c.Id == DocumentReaderZipRegistrationExtensions.HandlerId);
            Assert.Contains(result.Manifest.Handlers, c => c.Id == DocumentReaderHtmlRegistrationExtensions.HandlerId);
            Assert.Contains(result.Manifest.Handlers, c => c.Id == DocumentReaderJsonRegistrationExtensions.HandlerId);
            Assert.Contains(result.Manifest.Handlers, c => c.Id == DocumentReaderXmlRegistrationExtensions.HandlerId);

            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(result.ManifestJson));
            var jsonChunks = DocumentReaderJsonExtensions.ReadJson(
                stream,
                sourceName: "bootstrap-manifest.json",
                jsonOptions: new JsonReadOptions {
                    ChunkRows = 256,
                    IncludeMarkdown = false
                }).ToList();
            Assert.NotEmpty(jsonChunks);
            Assert.Contains(jsonChunks, c => (c.Text?.Contains("officeimo.reader.csv", StringComparison.Ordinal) ?? false));
        } finally {
            DocumentReaderCsvRegistrationExtensions.UnregisterCsvHandler();
            DocumentReaderEpubRegistrationExtensions.UnregisterEpubHandler();
            DocumentReaderZipRegistrationExtensions.UnregisterZipHandler();
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
            DocumentReaderJsonRegistrationExtensions.UnregisterJsonHandler();
            DocumentReaderXmlRegistrationExtensions.UnregisterXmlHandler();
        }
    }

    [Fact]
    public void DocumentReader_BootstrapHostFromLoadedAssemblies_NoMatches_ReturnsBuiltInManifest() {
        EnsureModularReaderAssembliesLoaded();

        try {
            DocumentReaderCsvRegistrationExtensions.UnregisterCsvHandler();
            DocumentReaderEpubRegistrationExtensions.UnregisterEpubHandler();
            DocumentReaderZipRegistrationExtensions.UnregisterZipHandler();
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
            DocumentReaderJsonRegistrationExtensions.UnregisterJsonHandler();
            DocumentReaderXmlRegistrationExtensions.UnregisterXmlHandler();

            var result = DocumentReader.BootstrapHostFromLoadedAssemblies(
                assemblyNamePrefix: "OfficeIMO.Reader.DoesNotExist.",
                options: new ReaderHostBootstrapOptions {
                    IncludeBuiltInCapabilities = true,
                    IncludeCustomCapabilities = true
                });

            Assert.NotNull(result);
            Assert.Equal("OfficeIMO.Reader.DoesNotExist.", result.AssemblyNamePrefix);
            Assert.Empty(result.RegisteredHandlers);
            Assert.Contains(result.Manifest.Handlers, c => c.IsBuiltIn && c.Id == "officeimo.reader.word");
            Assert.DoesNotContain(result.Manifest.Handlers, c => c.Id == DocumentReaderCsvRegistrationExtensions.HandlerId);
        } finally {
            DocumentReaderCsvRegistrationExtensions.UnregisterCsvHandler();
            DocumentReaderEpubRegistrationExtensions.UnregisterEpubHandler();
            DocumentReaderZipRegistrationExtensions.UnregisterZipHandler();
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
            DocumentReaderJsonRegistrationExtensions.UnregisterJsonHandler();
            DocumentReaderXmlRegistrationExtensions.UnregisterXmlHandler();
        }
    }

    [Fact]
    public void DocumentReader_BootstrapHostFromAssemblies_CanReturnCustomOnlyManifest() {
        try {
            DocumentReaderCsvRegistrationExtensions.UnregisterCsvHandler();
            DocumentReaderEpubRegistrationExtensions.UnregisterEpubHandler();
            DocumentReaderZipRegistrationExtensions.UnregisterZipHandler();
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
            DocumentReaderJsonRegistrationExtensions.UnregisterJsonHandler();
            DocumentReaderXmlRegistrationExtensions.UnregisterXmlHandler();

            var result = DocumentReader.BootstrapHostFromAssemblies(
                new[] {
                    typeof(DocumentReaderCsvRegistrationExtensions).Assembly,
                    typeof(DocumentReaderEpubRegistrationExtensions).Assembly,
                    typeof(DocumentReaderZipRegistrationExtensions).Assembly,
                    typeof(DocumentReaderHtmlRegistrationExtensions).Assembly,
                    typeof(DocumentReaderJsonRegistrationExtensions).Assembly,
                    typeof(DocumentReaderXmlRegistrationExtensions).Assembly
                },
                new ReaderHostBootstrapOptions {
                    ReplaceExistingHandlers = true,
                    IncludeBuiltInCapabilities = false,
                    IncludeCustomCapabilities = true,
                    IndentedManifestJson = true
                });

            Assert.NotNull(result);
            Assert.Null(result.AssemblyNamePrefix);
            Assert.True(result.ReplaceExistingHandlers);
            Assert.NotEmpty(result.RegisteredHandlers);
            Assert.DoesNotContain(result.Manifest.Handlers, c => c.IsBuiltIn);
            Assert.Contains(result.Manifest.Handlers, c => c.Id == DocumentReaderCsvRegistrationExtensions.HandlerId);
            Assert.Contains(Environment.NewLine, result.ManifestJson, StringComparison.Ordinal);
        } finally {
            DocumentReaderCsvRegistrationExtensions.UnregisterCsvHandler();
            DocumentReaderEpubRegistrationExtensions.UnregisterEpubHandler();
            DocumentReaderZipRegistrationExtensions.UnregisterZipHandler();
            DocumentReaderHtmlRegistrationExtensions.UnregisterHtmlHandler();
            DocumentReaderJsonRegistrationExtensions.UnregisterJsonHandler();
            DocumentReaderXmlRegistrationExtensions.UnregisterXmlHandler();
        }
    }

    [Theory]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("   ")]
    public void DocumentReader_BootstrapHostFromLoadedAssemblies_EmptyPrefix_Throws(string prefix) {
        Assert.Throws<ArgumentException>(() => DocumentReader.BootstrapHostFromLoadedAssemblies(prefix));
    }

    [Fact]
    public void DocumentReader_ModularRegistrationHelpers_DispatchesZipStream() {
        try {
            DocumentReaderZipRegistrationExtensions.RegisterZipHandler(replaceExisting: true);

            using var stream = new MemoryStream();
            using (var archive = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteTextEntry(archive, "docs/readme.md", "# Stream ZIP" + Environment.NewLine + Environment.NewLine + "Body from zip stream.");
            }

            stream.Position = 0;
            var chunks = DocumentReader.Read(stream, "bundle.zip").ToList();

            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Markdown &&
                (c.Location.Path?.Contains("bundle.zip::docs/readme.md", StringComparison.OrdinalIgnoreCase) ?? false) &&
                (c.Text?.Contains("Body from zip stream.", StringComparison.Ordinal) ?? false));
        } finally {
            DocumentReaderZipRegistrationExtensions.UnregisterZipHandler();
        }
    }

    [Fact]
    public void DocumentReader_ModularRegistrationHelpers_DispatchesZipNonSeekableStream() {
        try {
            DocumentReaderZipRegistrationExtensions.RegisterZipHandler(replaceExisting: true);

            using var zipBuffer = new MemoryStream();
            using (var archive = new ZipArchive(zipBuffer, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteTextEntry(archive, "docs/readme.md", "# Stream ZIP" + Environment.NewLine + Environment.NewLine + "Body from non-seekable stream.");
            }

            var bytes = zipBuffer.ToArray();
            using var stream = new NonSeekableReadStream(bytes);
            var chunks = DocumentReader.Read(stream, "bundle.zip").ToList();

            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Markdown &&
                (c.Location.Path?.Contains("bundle.zip::docs/readme.md", StringComparison.OrdinalIgnoreCase) ?? false) &&
                (c.Text?.Contains("Body from non-seekable stream.", StringComparison.Ordinal) ?? false));
        } finally {
            DocumentReaderZipRegistrationExtensions.UnregisterZipHandler();
        }
    }

    [Fact]
    public void DocumentReader_ModularRegistrationHelpers_DispatchesZipNonSeekableStream_EnforcesMaxInputBytes() {
        try {
            DocumentReaderZipRegistrationExtensions.RegisterZipHandler(replaceExisting: true);

            using var zipBuffer = new MemoryStream();
            using (var archive = new ZipArchive(zipBuffer, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteTextEntry(archive, "docs/readme.md", "# Stream ZIP" + Environment.NewLine + Environment.NewLine + "Body from non-seekable stream.");
            }

            var bytes = zipBuffer.ToArray();
            using var stream = new NonSeekableReadStream(bytes);
            var ex = Assert.Throws<IOException>(() => DocumentReader.Read(stream, "bundle.zip", new ReaderOptions { MaxInputBytes = 16 }).ToList());

            Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
        } finally {
            DocumentReaderZipRegistrationExtensions.UnregisterZipHandler();
        }
    }

    [Fact]
    public void DocumentReader_ModularRegistrationHelpers_DispatchesEpubStream() {
        try {
            DocumentReaderEpubRegistrationExtensions.RegisterEpubHandler(replaceExisting: true);

            var bytes = BuildSimpleEpubBytes();
            using var stream = new MemoryStream(bytes, writable: false);
            var chunks = DocumentReader.Read(stream, "book.epub").ToList();

            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Epub &&
                (c.Location.Path?.Contains("book.epub::OEBPS/chapter.xhtml", StringComparison.OrdinalIgnoreCase) ?? false) &&
                (c.Text?.Contains("EPUB stream body text.", StringComparison.Ordinal) ?? false));
        } finally {
            DocumentReaderEpubRegistrationExtensions.UnregisterEpubHandler();
        }
    }

    [Fact]
    public void DocumentReader_ModularRegistrationHelpers_DispatchesEpubNonSeekableStream() {
        try {
            DocumentReaderEpubRegistrationExtensions.RegisterEpubHandler(replaceExisting: true);

            var bytes = BuildSimpleEpubBytes();
            using var stream = new NonSeekableReadStream(bytes);
            var chunks = DocumentReader.Read(stream, "book.epub").ToList();

            Assert.Contains(chunks, c =>
                c.Kind == ReaderInputKind.Epub &&
                (c.Location.Path?.Contains("book.epub::OEBPS/chapter.xhtml", StringComparison.OrdinalIgnoreCase) ?? false) &&
                (c.Text?.Contains("EPUB stream body text.", StringComparison.Ordinal) ?? false));
        } finally {
            DocumentReaderEpubRegistrationExtensions.UnregisterEpubHandler();
        }
    }

    [Fact]
    public void DocumentReader_ModularRegistrationHelpers_DispatchesEpubNonSeekableStream_EnforcesMaxInputBytes() {
        try {
            DocumentReaderEpubRegistrationExtensions.RegisterEpubHandler(replaceExisting: true);

            var bytes = BuildSimpleEpubBytes();
            using var stream = new NonSeekableReadStream(bytes);
            var ex = Assert.Throws<IOException>(() => DocumentReader.Read(stream, "book.epub", new ReaderOptions { MaxInputBytes = 16 }).ToList());

            Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
        } finally {
            DocumentReaderEpubRegistrationExtensions.UnregisterEpubHandler();
        }
    }

    [Fact]
    public void DocumentReader_ModularRegistrationHelpers_DispatchesStructuredJsonStream() {
        try {
            DocumentReaderJsonRegistrationExtensions.RegisterJsonHandler(replaceExisting: true);

            var payload = "{\"service\":{\"name\":\"IX\",\"enabled\":true,\"ports\":[443,8443]}}";
            using var stream = new MemoryStream(Encoding.UTF8.GetBytes(payload), writable: false);
            var chunks = DocumentReader.Read(stream, "config.json").ToList();

            Assert.NotEmpty(chunks);
            Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Json, c.Kind));
            Assert.Contains(chunks, c =>
                (c.Location.Path?.Contains("config.json", StringComparison.OrdinalIgnoreCase) ?? false) &&
                (c.Text?.Contains("$.service.name", StringComparison.Ordinal) ?? false));
        } finally {
            DocumentReaderJsonRegistrationExtensions.UnregisterJsonHandler();
        }
    }

    [Fact]
    public void DocumentReader_ModularRegistrationHelpers_DispatchesStructuredCsvNonSeekableStream() {
        try {
            DocumentReaderCsvRegistrationExtensions.RegisterCsvHandler(replaceExisting: true);

            var payload = "Name,Role\nAlice,Admin\nBob,Ops\n";
            using var stream = new NonSeekableReadStream(Encoding.UTF8.GetBytes(payload));
            var chunks = DocumentReader.Read(stream, "users.csv").ToList();

            Assert.NotEmpty(chunks);
            Assert.All(chunks, c => Assert.Equal(ReaderInputKind.Csv, c.Kind));
            Assert.Contains(chunks, c =>
                string.Equals(c.Location.Path, "users.csv", StringComparison.OrdinalIgnoreCase) &&
                c.Tables != null &&
                c.Tables.Count > 0 &&
                c.Tables[0].Columns.Contains("Name", StringComparer.Ordinal));
        } finally {
            DocumentReaderCsvRegistrationExtensions.UnregisterCsvHandler();
        }
    }

    [Fact]
    public void DocumentReader_ModularRegistrationHelpers_DispatchesStructuredCsvNonSeekableStream_EnforcesMaxInputBytes() {
        try {
            DocumentReaderCsvRegistrationExtensions.RegisterCsvHandler(replaceExisting: true);

            var payload = "Name,Role\nAlice,Admin\nBob,Ops\n";
            using var stream = new NonSeekableReadStream(Encoding.UTF8.GetBytes(payload));
            var ex = Assert.Throws<IOException>(() => DocumentReader.Read(
                stream,
                "users.csv",
                new ReaderOptions { MaxInputBytes = 16 }).ToList());

            Assert.Contains("Input exceeds MaxInputBytes", ex.Message, StringComparison.Ordinal);
        } finally {
            DocumentReaderCsvRegistrationExtensions.UnregisterCsvHandler();
        }
    }

    private static byte[] BuildSimpleEpubBytes() {
        using var ms = new MemoryStream();
        using (var archive = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true)) {
            WriteTextEntry(archive, "mimetype", "application/epub+zip", CompressionLevel.NoCompression);
            WriteTextEntry(archive, "META-INF/container.xml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<container version=\"1.0\" xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\">" +
                "<rootfiles><rootfile full-path=\"OEBPS/content.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles>" +
                "</container>");

            WriteTextEntry(archive, "OEBPS/content.opf",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<package version=\"3.0\" xmlns=\"http://www.idpf.org/2007/opf\">" +
                "<manifest><item id=\"ch1\" href=\"chapter.xhtml\" media-type=\"application/xhtml+xml\"/></manifest>" +
                "<spine><itemref idref=\"ch1\"/></spine>" +
                "</package>");

            WriteTextEntry(archive, "OEBPS/chapter.xhtml",
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?>" +
                "<html xmlns=\"http://www.w3.org/1999/xhtml\"><body><p>EPUB stream body text.</p></body></html>");
        }

        return ms.ToArray();
    }

    private static void WriteTextEntry(ZipArchive archive, string path, string content, CompressionLevel compressionLevel = CompressionLevel.Optimal) {
        var entry = archive.CreateEntry(path, compressionLevel);
        using var stream = entry.Open();
        using var writer = new StreamWriter(stream, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false), 4096, leaveOpen: false);
        writer.Write(content);
    }

    private static void EnsureModularReaderAssembliesLoaded() {
        _ = typeof(DocumentReaderCsvRegistrationExtensions);
        _ = typeof(DocumentReaderEpubRegistrationExtensions);
        _ = typeof(DocumentReaderZipRegistrationExtensions);
        _ = typeof(DocumentReaderHtmlRegistrationExtensions);
        _ = typeof(DocumentReaderJsonRegistrationExtensions);
        _ = typeof(DocumentReaderXmlRegistrationExtensions);
        _ = typeof(DocumentReaderTextRegistrationExtensions);
    }
}
