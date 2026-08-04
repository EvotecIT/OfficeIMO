using OfficeIMO.Email.Store;
using OfficeIMO.Drawing;
using OfficeIMO.Reader;
using OfficeIMO.Reader.All;
using OfficeIMO.Reader.Email;
using OfficeIMO.Tool.Commands.Reader;

namespace OfficeIMO.Tool.Agent;

internal sealed partial class OfficeImoAgentService {
    internal const int DefaultInspectOutputCharacters = 2_000;
    internal const int DefaultSearchOutputCharacters = 6_000;
    internal const int DefaultFetchOutputCharacters = 8_000;
    internal const int DefaultCapabilitiesOutputCharacters = 4_000;
    internal const int MinimumOutputCharacters = 512;
    internal const int MaximumOutputCharacters = 64_000;
    internal const int MaximumSearchResults = 25;
    internal const int MaximumSearchCursor = 10_000;

    private static readonly EmailStoreItemReadParts AgentEmailParts =
        EmailStoreItemReadParts.Metadata |
        EmailStoreItemReadParts.Bodies |
        EmailStoreItemReadParts.Recipients |
        EmailStoreItemReadParts.AttachmentMetadata;

    private readonly AgentPathPolicy _pathPolicy;
    private readonly AgentSourceRegistry _registry;

    internal OfficeImoAgentService(
        AgentPathPolicy? pathPolicy = null,
        AgentSourceRegistry? registry = null) {
        _pathPolicy = pathPolicy ?? AgentPathPolicy.FromEnvironment();
        _registry = registry ?? new AgentSourceRegistry();
    }

    internal async Task<AgentInspectResult> InspectAsync(
        string path,
        int maxOutputCharacters = DefaultInspectOutputCharacters,
        CancellationToken cancellationToken = default) {
        maxOutputCharacters = ValidateOutputBudget(maxOutputCharacters);
        string inputPath = _pathPolicy.ResolveInput(path);
        AgentSourceRegistration source = _registry.Register(inputPath, cancellationToken);
        AgentInspectResult result = IsEmailStoreSource(source)
            ? InspectEmailStore(source, maxOutputCharacters, cancellationToken)
            : CreateDocumentInspectResult(
                source,
                await ReadDocumentAsync(inputPath, cancellationToken).ConfigureAwait(false));
        TrimInspect(result, maxOutputCharacters);
        return result;
    }

    internal async Task<AgentSearchResult> SearchAsync(
        string path,
        string? query = null,
        string? subject = null,
        string? sender = null,
        string? folderId = null,
        DateTimeOffset? since = null,
        DateTimeOffset? before = null,
        bool? hasAttachments = null,
        bool? isRead = null,
        bool includeDescendants = false,
        int take = 10,
        int cursor = 0,
        int maxOutputCharacters = DefaultSearchOutputCharacters,
        CancellationToken cancellationToken = default) {
        ValidateSearchBounds(take, cursor);
        maxOutputCharacters = ValidateOutputBudget(maxOutputCharacters);
        string inputPath = _pathPolicy.ResolveInput(path);
        AgentSourceRegistration source = _registry.Register(inputPath, cancellationToken);
        AgentSearchResult result;
        if (IsEmailStoreSource(source)) {
            result = SearchEmailStore(
                source,
                query,
                subject,
                sender,
                folderId,
                since,
                before,
                hasAttachments,
                isRead,
                includeDescendants,
                take,
                cursor,
                cancellationToken);
        } else {
            if (string.IsNullOrWhiteSpace(query)) {
                throw new AgentUsageException("Document search requires a non-empty query.");
            }
            if (subject != null || sender != null || folderId != null || since.HasValue ||
                before.HasValue || hasAttachments.HasValue || isRead.HasValue || includeDescendants) {
                throw new AgentUsageException(
                    "Email-store filters are only valid for PST, OST, OLM, EMLX, Mbox, MBX, or mailbox directories.");
            }
            OfficeDocumentReadResult document =
                await ReadDocumentAsync(inputPath, cancellationToken).ConfigureAwait(false);
            result = SearchDocument(source, document, query!, take, cursor);
        }
        TrimSearch(result, maxOutputCharacters, cursor);
        return result;
    }

    internal async Task<AgentFetchResult> FetchAsync(
        string sourceId,
        string id,
        int cursor = 0,
        int maxOutputCharacters = DefaultFetchOutputCharacters,
        string? sourcePath = null,
        CancellationToken cancellationToken = default) {
        if (cursor < 0) throw new AgentUsageException("Fetch cursor must not be negative.");
        maxOutputCharacters = ValidateOutputBudget(maxOutputCharacters);
        AgentSourceRegistration source = string.IsNullOrWhiteSpace(sourcePath)
            ? _registry.Resolve(sourceId, cancellationToken)
            : _registry.Resolve(
                sourceId,
                _pathPolicy.ResolveInput(sourcePath),
                cancellationToken);
        (string kind, string value) = AgentOpaqueId.Decode(id);
        AgentFetchResult result;
        if (kind == "mail") {
            if (!IsEmailStoreSource(source)) {
                throw new AgentUsageException("The mail result id does not belong to an email store.");
            }
            result = FetchEmailStoreItem(source, value, id, cursor, cancellationToken);
        } else {
            if (IsEmailStoreSource(source)) {
                throw new AgentUsageException("Email-store fetch requires a mail result id.");
            }
            OfficeDocumentReadResult document =
                await ReadDocumentAsync(source.Path, cancellationToken).ConfigureAwait(false);
            result = FetchDocument(source, document, kind, value, id, cursor);
        }
        TrimFetch(result, maxOutputCharacters, cursor);
        return result;
    }

    internal AgentCapabilitiesResult Capabilities(
        string? extension = null,
        string operation = "read",
        int maxOutputCharacters = DefaultCapabilitiesOutputCharacters) {
        maxOutputCharacters = ValidateOutputBudget(maxOutputCharacters);
        operation = NormalizeOperation(operation);
        string? normalizedExtension = NormalizeExtension(extension);
        OfficeDocumentReader reader = CreateReader();
        var capabilities = reader.GetCapabilities()
            .Where(capability => normalizedExtension == null ||
                capability.Extensions.Contains(normalizedExtension, StringComparer.OrdinalIgnoreCase))
            .Where(capability =>
                operation != "convert" ||
                (capability.Id != OfficeDocumentReaderBuilderEmailStoreExtensions.HandlerId &&
                 capability.Id != OfficeDocumentReaderBuilderEmailExtensions.MailboxHandlerId))
            .OrderBy(capability => capability.Id, StringComparer.Ordinal)
            .Select(capability => new AgentCapabilitySummary {
                Id = capability.Id,
                Name = capability.DisplayName,
                Kind = capability.Kind.ToString(),
                Extensions = capability.Extensions
                    .OrderBy(value => value, StringComparer.OrdinalIgnoreCase)
                    .ToArray()
            })
            .ToList();
        var conversions = operation == "convert"
            ? OfficeConversionCapabilityCatalog.AgentRoutes
                .Where(route => normalizedExtension == null ||
                    route.SourceExtensions.Contains(normalizedExtension, StringComparer.OrdinalIgnoreCase))
                .Select(route => new AgentConversionCapabilitySummary {
                    Id = route.Id,
                    Source = route.Source,
                    Target = route.Target,
                    SourceExtensions = route.SourceExtensions,
                    TargetExtension = route.TargetExtension,
                    PackageId = route.PackageId,
                    Fidelity = route.Fidelity.ToString(),
                    ResultContract = route.ResultContract,
                    BrowserAvailable = route.BrowserAvailable
                })
                .ToList()
            : new List<AgentConversionCapabilitySummary>();
        var result = new AgentCapabilitiesResult {
            Extension = normalizedExtension,
            Operation = operation,
            Returned = capabilities.Count,
            Capabilities = capabilities,
            ConversionReturned = conversions.Count,
            Conversions = conversions
        };
        while (AgentJson.Measure(result) > maxOutputCharacters && (conversions.Count > 0 || capabilities.Count > 0)) {
            if (conversions.Count > 0) {
                conversions.RemoveAt(conversions.Count - 1);
                result.ConversionReturned = conversions.Count;
            } else {
                capabilities.RemoveAt(capabilities.Count - 1);
                result.Returned = capabilities.Count;
            }
            result.Truncated = true;
        }
        return result;
    }

    internal async Task<AgentConvertResult> ConvertAsync(
        string path,
        string outputPath,
        string format = "markdown",
        bool overwrite = false,
        CancellationToken cancellationToken = default) {
        string inputPath = _pathPolicy.ResolveInput(path);
        AgentSourceRegistration source = _registry.Register(inputPath, cancellationToken);
        if (IsEmailStoreSource(source)) {
            throw new AgentUsageException(
                "Whole-store conversion is intentionally disabled. Search the store and fetch selected items.");
        }
        string destination = _pathPolicy.ResolveOutput(outputPath);
        ReaderToolPathSafety.EnsureDistinctFile(inputPath, destination);
        if (!overwrite && File.Exists(destination)) {
            throw new AgentUsageException(
                "Output file already exists. Choose a new path or explicitly enable overwrite.");
        }
        ReaderToolOutputFormat outputFormat = format.ToLowerInvariant() switch {
            "markdown" or "md" => ReaderToolOutputFormat.Markdown,
            "json" => ReaderToolOutputFormat.Json,
            _ => throw new AgentUsageException("Conversion format must be 'markdown' or 'json'.")
        };
        OfficeDocumentReadResult document =
            await ReadDocumentAsync(inputPath, cancellationToken).ConfigureAwait(false);
        string content = ReaderToolOutput.FormatDocument(document, outputFormat);
        await ReaderToolOutput.WriteFileAsync(
            destination,
            content,
            overwrite,
            cancellationToken).ConfigureAwait(false);
        string hash;
        await using (FileStream stream = new FileStream(
            destination,
            FileMode.Open,
            FileAccess.Read,
            FileShare.Read,
            64 * 1024,
            useAsync: true)) {
            byte[] bytes = await System.Security.Cryptography.SHA256.HashDataAsync(
                stream,
                cancellationToken).ConfigureAwait(false);
            hash = Convert.ToHexString(bytes).ToLowerInvariant();
        }
        return new AgentConvertResult {
            SourceId = source.SourceId,
            SourcePath = source.Path,
            OutputPath = destination,
            Format = outputFormat == ReaderToolOutputFormat.Json ? "json" : "markdown",
            LengthBytes = new FileInfo(destination).Length,
            Sha256 = hash,
            DiagnosticCount = document.Diagnostics.Count
        };
    }

    private static OfficeDocumentReader CreateReader() {
        var email = new ReaderEmailHandlersOptions {
            Artifacts = new ReaderEmailOptions {
                IncludeAttachmentContent = false
            },
            Stores = new ReaderEmailStoreOptions {
                ItemReadOptions = new EmailStoreItemReadOptions(
                    AgentEmailParts,
                    preferStreamingAttachmentContent: true),
                MaxItems = MaximumSearchResults,
                StreamAttachmentContent = true,
                ComputeSourceHash = false
            }
        };
        return new OfficeDocumentReaderBuilder()
            .AddAllOfficeIMOHandlers(new ReaderAllOptions { Email = email })
            .Build();
    }

    private static async Task<OfficeDocumentReadResult> ReadDocumentAsync(
        string path,
        CancellationToken cancellationToken) {
        OfficeDocumentReader reader = CreateReader();
        return await reader.ReadDocumentAsync(
            path,
            new ReaderOptions {
                MaxInputBytes = ReaderToolArguments.DefaultMaxInputBytes,
                MaxChars = DefaultFetchOutputCharacters,
                ComputeHashes = false
            },
            cancellationToken).ConfigureAwait(false);
    }

    private static bool IsEmailStoreSource(AgentSourceRegistration source) {
        if (source.IsDirectory) return true;
        string extension = Path.GetExtension(source.Path);
        return extension.Equals(".pst", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".ost", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".olm", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".emlx", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".mbox", StringComparison.OrdinalIgnoreCase) ||
               extension.Equals(".mbx", StringComparison.OrdinalIgnoreCase);
    }

    private static int ValidateOutputBudget(int value) {
        if (value < MinimumOutputCharacters || value > MaximumOutputCharacters) {
            throw new AgentUsageException(
                "Maximum output characters must be between " +
                MinimumOutputCharacters + " and " + MaximumOutputCharacters + ".");
        }
        return value;
    }

    private static void ValidateSearchBounds(int take, int cursor) {
        if (take < 1 || take > MaximumSearchResults) {
            throw new AgentUsageException(
                "Search take must be between 1 and " + MaximumSearchResults + ".");
        }
        if (cursor < 0 || cursor > MaximumSearchCursor) {
            throw new AgentUsageException(
                "Search cursor must be between 0 and " + MaximumSearchCursor + ".");
        }
    }

    private static string NormalizeOperation(string operation) {
        string normalized = string.IsNullOrWhiteSpace(operation)
            ? "read"
            : operation.Trim().ToLowerInvariant();
        return normalized switch {
            "read" or "inspect" or "search" or "fetch" or "convert" => normalized,
            _ => throw new AgentUsageException(
                "Capability operation must be read, inspect, search, fetch, or convert.")
        };
    }

    private static string? NormalizeExtension(string? extension) {
        if (string.IsNullOrWhiteSpace(extension)) return null;
        string value = extension.Trim();
        string normalized = value.StartsWith(".", StringComparison.Ordinal) ? value : "." + value;
        if (normalized.Length > 64) {
            throw new AgentUsageException("Capability extensions must not exceed 64 characters.");
        }
        return normalized;
    }
}
