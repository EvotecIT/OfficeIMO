using System;
using System.IO;
using Xunit;

namespace OfficeIMO.Xps.Tests;

public sealed class NativeSaveLifecycleTests {
    [Fact]
    public void SaveBindsOnlySuccessfulPathsAndLeavesStreamsOpen() {
        var document = XpsDocument.Create();
        Assert.Throws<InvalidOperationException>(() => document.Save());
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-Xps-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string first = Path.Combine(directory, "first.xps");
            document.Save(first);
            Assert.Equal(Path.GetFullPath(first), document.FilePath);
            Assert.ThrowsAny<IOException>(() => document.Save(Path.Combine(directory, "missing", "file.xps")));
            Assert.Equal(Path.GetFullPath(first), document.FilePath);
            using var output = new MemoryStream();
            document.Save(output);
            Assert.True(output.CanWrite);
            Assert.Equal(Path.GetFullPath(first), document.FilePath);
            Assert.Equal(document.ToBytes(), output.ToArray());
            var reopened = XpsDocument.Load(first);
            Assert.Equal(Path.GetFullPath(first), reopened.FilePath);
            reopened.Save();
            Assert.NotEmpty(File.ReadAllBytes(first));
        } finally { Directory.Delete(directory, recursive: true); }
    }
}
