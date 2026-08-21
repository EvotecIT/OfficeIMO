using System.IO;
using Xunit;

namespace OfficeIMO.Provenance.C2pa.Tests;

public sealed partial class C2paToolProvenanceVerifierTestsWave72 {
    [Fact]
    public void ProcessRunnerRejectsNonUtf8ProviderOutput() {
        using var output = new MemoryStream(new byte[] { 0xFF });

        Assert.Throws<InvalidDataException>(() => C2paToolProcessRunner
            .ReadBoundedAsync(output, 1024, "standard output")
            .GetAwaiter()
            .GetResult());
    }
}
