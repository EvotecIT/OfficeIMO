using OfficeIMO.Publisher.Internal;

namespace OfficeIMO.Publisher;

public sealed partial class PublisherDocument {
    internal PublisherQuillStyleCatalog SourceStyleCatalog => _model?.SourceStyleCatalog ?? PublisherQuillStyleCatalog.Empty;
    /// <summary>Prepares an opt-in internal candidate without publishing output or changing the model's source baseline.</summary>
    internal PublisherNativePreparation PrepareNativeBytes(long maximumBytes, CancellationToken token = default) =>
        _model?.PrepareNativeBytes(maximumBytes, token) ?? throw new NotSupportedException("The publication has no semantic model.");
}
