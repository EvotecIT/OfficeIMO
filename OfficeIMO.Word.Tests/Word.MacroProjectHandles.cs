using System;
using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class WordMacroProjectHandleTests {
    [Fact]
    public void ProjectHandlesReadTheAttachedBinaryAndRejectReplacement() {
        using var document = WordDocument.Create();
        // Opaque projects are accepted without introducing eager source parsing.
        byte[] original = { 1, 2, 3, 4 };
        var project = document.AddMacro(original);
        Assert.Same(document, project.Document);
        Assert.EndsWith("vbaProject.bin", project.PartUri);
        Assert.Equal(original, project.ToBytes());
        Assert.Throws<InvalidDataException>(() => project.ToBytes(2));
        using var replacement = new MemoryStream(new byte[] { 5, 6 });
        var current = document.AddMacro(replacement);
        Assert.Throws<InvalidOperationException>(() => project.Remove());
        Assert.Throws<InvalidOperationException>(() => project.ToBytes());
        Assert.Equal(new byte[] { 5, 6 }, current.ToBytes());
        current.Remove();
        Assert.False(document.HasMacros);
        Assert.Throws<InvalidOperationException>(() => current.ToBytes());
    }
}
