using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Runs blocking OCR provider fixtures apart from other test collections.</summary>
[CollectionDefinition(Name, DisableParallelization = true)]
public sealed class ReaderOcrCoreCollection {
    public const string Name = "Reader OCR core providers";
}
