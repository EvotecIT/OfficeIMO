using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives.Tests;

public sealed class ArchiveCompoundQuotaPropagationTests {
    [Theory]
    [InlineData("full", true)]
    [InlineData("selective", true)]
    [InlineData("view", true)]
    [InlineData("directory", true)]
    [InlineData("full", false)]
    [InlineData("selective", false)]
    [InlineData("view", false)]
    public void CanonicalSiblingPathsPreserveTerminalQuotaIdentityAndViewOwnership(string route, bool directory) {
        byte[] bytes = ArchiveFixtures.Msi();
        var options = directory ? new OfficeCompoundReadOptions(maxDirectoryEntries: 3, preserveResourceLimits: true)
            : new OfficeCompoundReadOptions(maxStreamBytes: 1, preserveResourceLimits: true);
        using var source = new MemoryStream(new byte[7].Concat(bytes).ToArray()); source.Position = 7;
        OfficeCompoundStreamLimitExceededException error = Assert.Throws<OfficeCompoundStreamLimitExceededException>(() => {
            switch (route) {
                case "full": OfficeCompoundFileReader.TryRead(bytes, options, out OfficeCompoundFile? _, out _); break;
                case "selective": OfficeCompoundFileReader.TryReadSelective(source, options, (_, _) => false,
                    (_, _) => throw new InvalidOperationException("No destination is needed."), out _, out _); break;
                case "view": OfficeCompoundFileReader.TryOpenStream(source, options, (_, _) => true, true, default, out _, out _); break;
                case "directory": OfficeCompoundFileReader.TryInspectDirectory(source, long.MaxValue, 3, out _, out _, preserveResourceLimits: true); break;
            }
        });
        Assert.Equal(directory ? "MaxDirectoryEntries" : "MaxStreamBytes", error.LimitName);
        Assert.Equal(directory ? 3 : 1, error.MaximumValue);
        Assert.True(error.ActualValue > error.MaximumValue); Assert.True(source.CanRead);
        if (route is "selective" or "view") Assert.Equal(7, source.Position);
    }

    [Theory]
    [InlineData("full")]
    [InlineData("selective")]
    [InlineData("view")]
    [InlineData("directory")]
    public void NativeTryPathsStillReturnMalformedProfilesWithoutAQuotaException(string route) {
        byte[] bytes = new byte[1024];
        using var source = new MemoryStream(bytes);
        var options = new OfficeCompoundReadOptions();
        bool result = route switch {
            "full" => OfficeCompoundFileReader.TryRead(bytes, options, out OfficeCompoundFile? _, out _),
            "selective" => OfficeCompoundFileReader.TryReadSelective(source, options, (_, _) => false,
                (_, _) => throw new InvalidOperationException("No destination is needed."), out _, out _),
            "view" => OfficeCompoundFileReader.TryOpenStream(source, options, (_, _) => true, true, default, out _, out _),
            _ => OfficeCompoundFileReader.TryInspectDirectory(source, long.MaxValue, 65536, out _, out _)
        };
        Assert.False(result); Assert.True(source.CanRead);
    }
}
