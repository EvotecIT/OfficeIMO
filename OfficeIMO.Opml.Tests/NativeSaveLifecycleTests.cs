using System;
using System.IO;
using Xunit;

namespace OfficeIMO.Opml.Tests;

public sealed class NativeSaveLifecycleTests {
    [Fact]
    public void SaveBindsOnlySuccessfulPathsAndLeavesStreamsOpen() {
        var document = OpmlDocument.Create();
        Assert.Throws<InvalidOperationException>(() => document.Save());
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-Opml-" + Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        try {
            string first = Path.Combine(directory, "first.opml");
            document.Save(first);
            Assert.Equal(Path.GetFullPath(first), document.FilePath);
            string blocked = Path.Combine(directory, "blocked");
            File.WriteAllText(blocked, "occupied");
            Assert.ThrowsAny<IOException>(() => document.Save(Path.Combine(blocked, "file.opml")));
            Assert.Equal(Path.GetFullPath(first), document.FilePath);
            using var output = new MemoryStream();
            document.Save(output);
            Assert.True(output.CanWrite);
            Assert.Equal(Path.GetFullPath(first), document.FilePath);
            Assert.Equal(document.ToBytes(), output.ToArray());
            var reopened = OpmlDocument.Load(first);
            Assert.Equal(Path.GetFullPath(first), reopened.FilePath);
            reopened.Save();
            Assert.NotEmpty(File.ReadAllBytes(first));
        } finally { Directory.Delete(directory, recursive: true); }
    }
}
