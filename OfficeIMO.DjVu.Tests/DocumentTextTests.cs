using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using Xunit;

namespace OfficeIMO.DjVu.Tests;

public sealed class DocumentTextTests {
    private static string Fixture(string name) => Path.Combine(AppContext.BaseDirectory, "Fixtures", "Archive", name);

    [Fact]
    public void ArchivalPageRecoversStoredTextAndWordGeometry() {
        var document = DjVuDocument.Load(Fixture("time-machine-16.djvu"));
        var page = Assert.Single(document.Pages);
        Assert.Equal((1552, 2826, 400, 0), (page.Width, page.Height, page.Dpi, page.Rotation));
        var result = page.GetText();
        Assert.Equal(DjVuTextStatus.Present, result.Status);
        Assert.Equal(File.ReadAllText(Fixture("time-machine-16.txt"), Encoding.UTF8), result.Text);
        Assert.Null(result.Diagnostic);
        var word = Flatten(result.Zones).First(z => z.Kind == DjVuTextZoneKind.Word);
        Assert.Equal("4 ", result.Text.Substring(word.CharacterOffset, word.CharacterLength));
        Assert.Equal((201, 2627, 28, 36), (word.Bounds.X, word.Bounds.Y, word.Bounds.Width, word.Bounds.Height));
    }

    [Fact]
    public void BundledDirectoryPreservesOrderAndMissingText() {
        var document = DjVuDocument.Load(Fixture("time-machine-selected.djvu"));
        Assert.Equal(2, document.Pages.Count);
        Assert.Equal(new[] { 1, 2 }, document.Pages.Select(p => p.Number));
        Assert.Equal(DjVuTextStatus.Absent, document.Pages[0].GetText().Status);
        Assert.Equal(File.ReadAllText(Fixture("time-machine-16.txt")), document.Pages[1].GetText().Text);
        Assert.NotEqual(document.Pages[0].Id, document.Pages[1].Id);
    }

    [Fact]
    public void InputAndOptionsAreOwnedAndStreamStartsAtCurrentPosition() {
        byte[] original = File.ReadAllBytes(Fixture("time-machine-selected.djvu"));
        byte[] bytes = new byte[original.Length + 7];
        original.CopyTo(bytes, 7);
        var options = new DjVuReadOptions();
        using var stream = new MemoryStream(bytes);
        stream.Position = 7;
        var document = DjVuDocument.Load(stream, options);
        Assert.True(stream.CanRead);
        Assert.Equal(stream.Length, stream.Position);
        Array.Clear(bytes, 0, bytes.Length);
        options.MaxTextCharacters = 1;
        Assert.Equal(DjVuTextStatus.Present, document.Pages[1].GetText().Status);
        Assert.Equal(document.SourceSha256, DjVuDocument.Load(original).SourceSha256);
    }

    [Fact]
    public void ExpansionAndSourceLimitsAndCancellationAreConsequential() {
        byte[] bytes = File.ReadAllBytes(Fixture("time-machine-selected.djvu"));
        Assert.Equal(nameof(DjVuReadOptions.MaxSourceBytes), Assert.Throws<DjVuResourceLimitException>(() => DjVuDocument.Load(bytes, new DjVuReadOptions { MaxSourceBytes = bytes.Length - 1 })).LimitName);
        Assert.Equal(nameof(DjVuReadOptions.MaxExpandedBytes), Assert.Throws<DjVuResourceLimitException>(() => DjVuDocument.Load(bytes, new DjVuReadOptions { MaxExpandedBytes = 8 })).LimitName);
        Assert.Equal(nameof(DjVuReadOptions.MaxPages), Assert.Throws<DjVuResourceLimitException>(() => DjVuDocument.Load(bytes, new DjVuReadOptions { MaxPages = 1 })).LimitName);
        Assert.Throws<OperationCanceledException>(() => DjVuDocument.Load(bytes, cancellationToken: new CancellationToken(true)));
    }

    private static IEnumerable<DjVuTextZone> Flatten(IEnumerable<DjVuTextZone> zones) {
        foreach (var zone in zones) {
            yield return zone;
            foreach (var child in Flatten(zone.Children)) yield return child;
        }
    }
}
