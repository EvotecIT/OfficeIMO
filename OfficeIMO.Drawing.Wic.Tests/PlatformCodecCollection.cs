using Xunit;

namespace OfficeIMO.Drawing.Wic.Tests;

// Registration mutates a process-wide Core fallback. Never overlap it with another test collection.
[CollectionDefinition("WIC platform registration", DisableParallelization = true)]
public sealed class PlatformCodecCollection { }
