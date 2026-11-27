using Microsoft.Extensions.DocumentExtraction;
using OfficeIMO.Reader;
using OfficeIMO.Reader.DocumentExtraction;

namespace Microsoft.Extensions.DependencyInjection;

/// <summary>Registers <see cref="OfficeIMODocumentExtractionClient"/> as a <see cref="DocumentExtractionClient"/>.</summary>
public static class OfficeIMODocumentExtractionServiceCollectionExtensions {
    /// <summary>
    /// Registers the OfficeIMO client as the default <see cref="DocumentExtractionClient"/>. The returned builder adds
    /// middleware such as <c>UseLogging()</c>.
    /// </summary>
    /// <param name="services">The service collection.</param>
    /// <param name="configureReader">
    /// Registers reader handlers, for example <c>b =&gt; b.AddAllOfficeIMOHandlers()</c>. When null, an
    /// <see cref="OfficeDocumentReader"/> registered in the container is used, or <see cref="OfficeDocumentReader.Default"/>.
    /// </param>
    /// <param name="configureOptions">Optional client defaults.</param>
    /// <param name="lifetime">The client lifetime. The reader is thread-safe, so singleton is the default.</param>
    public static DocumentExtractionClientBuilder AddOfficeIMODocumentExtraction(
        this IServiceCollection services,
        Action<OfficeDocumentReaderBuilder>? configureReader = null,
        Action<OfficeIMODocumentExtractionOptions>? configureOptions = null,
        ServiceLifetime lifetime = ServiceLifetime.Singleton) {
        ArgumentNullException.ThrowIfNull(services);
        return services.AddDocumentExtractionClient(CreateFactory(configureReader, configureOptions), lifetime);
    }

    /// <summary>
    /// Registers the OfficeIMO client as a keyed <see cref="DocumentExtractionClient"/>, so it can sit next to other
    /// providers and be resolved with <c>[FromKeyedServices]</c>.
    /// </summary>
    /// <param name="services">The service collection.</param>
    /// <param name="serviceKey">The registration key.</param>
    /// <param name="configureReader">Registers reader handlers. See <see cref="AddOfficeIMODocumentExtraction"/>.</param>
    /// <param name="configureOptions">Optional client defaults.</param>
    /// <param name="lifetime">The client lifetime.</param>
    public static DocumentExtractionClientBuilder AddKeyedOfficeIMODocumentExtraction(
        this IServiceCollection services,
        object? serviceKey,
        Action<OfficeDocumentReaderBuilder>? configureReader = null,
        Action<OfficeIMODocumentExtractionOptions>? configureOptions = null,
        ServiceLifetime lifetime = ServiceLifetime.Singleton) {
        ArgumentNullException.ThrowIfNull(services);
        return services.AddKeyedDocumentExtractionClient(serviceKey, CreateFactory(configureReader, configureOptions), lifetime);
    }

    private static Func<IServiceProvider, DocumentExtractionClient> CreateFactory(
        Action<OfficeDocumentReaderBuilder>? configureReader,
        Action<OfficeIMODocumentExtractionOptions>? configureOptions) {
        var options = new OfficeIMODocumentExtractionOptions();
        configureOptions?.Invoke(options);
        // One reader per registration: building it in the factory keeps handler setup out of service resolution paths.
        Lazy<OfficeDocumentReader>? configured = configureReader == null
            ? null
            : new Lazy<OfficeDocumentReader>(() => {
                var builder = new OfficeDocumentReaderBuilder();
                configureReader(builder);
                return builder.Build();
            });
        return services => {
            OfficeDocumentReader reader = configured?.Value
                ?? (OfficeDocumentReader?)services.GetService(typeof(OfficeDocumentReader))
                ?? OfficeDocumentReader.Default;
            return new OfficeIMODocumentExtractionClient(reader, options);
        };
    }
}
