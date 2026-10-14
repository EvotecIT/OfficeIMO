using Xunit;

namespace OfficeIMO.Tests.Pdf;

// These contracts count preparation in the bounded process-wide content cache.
// Unrelated image tests can legitimately evict its entries between two writes.
// Tests inside the collection retain their own concurrent-document scenarios.
[CollectionDefinition(Name, DisableParallelization = true)]
public sealed class PdfImagePreparationCacheCollection {
    public const string Name = "PDF image preparation cache";
}
