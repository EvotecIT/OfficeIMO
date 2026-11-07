using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingFontSnapshotTemplateTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImmutableSnapshotAdmissionPreservesColdAndWarmBudgetFailures(bool cff) {
        byte[] data = FontData(cff);
        int retainedBytes = checked(data.Length * (cff ? 3 : 2));
        IOfficeFontProgram? firstProgram = null;
        for (int pass = 0; pass < 2; pass++) {
            foreach (int limit in new[] { data.Length - 1, retainedBytes - 1, retainedBytes, retainedBytes + 1 }) {
                var expected = new OfficeFontFaceCollection();
                bool expectedAccepted = expected.TryAddBounded("Snapshot", data, OfficeFontFaceDescriptor.Regular,
                    OfficeFontUnicodeRangeSet.All, limit, out int expectedBytes, out string? expectedError);
                var actual = new OfficeFontFaceCollection();
                bool accepted = actual.TryAddImmutableSnapshotBounded("Snapshot", data, OfficeFontFaceDescriptor.Regular,
                    limit, out int decodedBytes, out string? error);

                Assert.Equal(limit >= retainedBytes, accepted);
                Assert.Equal(expectedAccepted, accepted);
                Assert.Equal(expectedBytes, decodedBytes);
                Assert.Equal(expectedError, error);
                Assert.Equal(limit < data.Length
                    ? "OpenType font data exceeds the configured decoded byte limit."
                    : limit < retainedBytes ? "Decoded font data exceeds the configured byte limit." : null, error);
                if (accepted) {
                    Assert.Equal(retainedBytes, decodedBytes);
                    OfficeFontFace face = Assert.Single(actual.Faces);
                    if (firstProgram == null) firstProgram = face.Program;
                    else Assert.Same(firstProgram, face.Program);
                } else {
                    Assert.Equal(0, decodedBytes);
                    Assert.Empty(actual.Faces);
                }
            }
        }
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeFontFaceCollection()
            .TryAddImmutableSnapshotBounded("Snapshot", data, OfficeFontFaceDescriptor.Regular, 0, out _, out _));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public async Task ImmutableSnapshotAdmissionSharesNativeProgramsAcrossConcurrentCollections(bool cff) {
        byte[] data = FontData(cff);
        int retainedBytes = checked(data.Length * (cff ? 3 : 2));
        string[] texts = { "A", "B", "AB", "BA" };
        OfficeFontFace[] references = Enumerable.Range(0, texts.Length).Select(index =>
            Assert.Single(new OfficeFontFaceCollection().Add("Alias " + index, data,
                OfficeFontFaceDescriptor.FromStyle(index % 2 == 0 ? OfficeFontStyle.Regular : OfficeFontStyle.Bold)).Faces)).ToArray();
        var ready = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        var start = new TaskCompletionSource<bool>(TaskCreationOptions.RunContinuationsAsynchronously);
        int waiting = 0;
        Task<OfficeFontFace>[] calls = Enumerable.Range(0, texts.Length).Select(index => Task.Run(async () => {
            if (Interlocked.Increment(ref waiting) == texts.Length) ready.SetResult(true);
            await start.Task;
            var fonts = new OfficeFontFaceCollection();
            OfficeFontFace reference = references[index];
            Assert.True(fonts.TryAddImmutableSnapshotBounded(reference.FamilyName, data, reference.Descriptor,
                retainedBytes, out int decodedBytes, out string? error), error);
            Assert.Equal(retainedBytes, decodedBytes);
            OfficeFontFace face = Assert.Single(fonts.Faces);
            AssertEquivalentFontUse(reference, face, texts[index]);
            return face;
        })).ToArray();
        await ready.Task;
        start.SetResult(true);
        OfficeFontFace[] faces = await Task.WhenAll(calls);
        foreach (OfficeFontFace face in faces) Assert.Same(faces[0].Program, face.Program);

        var warm = new OfficeFontFaceCollection();
        Assert.True(warm.TryAddImmutableSnapshotBounded(references[0].FamilyName, data, references[0].Descriptor,
            retainedBytes, out int warmBytes, out string? warmError), warmError);
        Assert.Equal(retainedBytes, warmBytes);
        OfficeFontFace warmFace = Assert.Single(warm.Faces);
        Assert.Same(faces[0].Program, warmFace.Program);
        AssertEquivalentFontUse(references[0], warmFace, texts[0]);
    }

    [Theory]
    [InlineData("ProviderOnly")]
    [InlineData("ResolverOnly")]
    public void ImmutableSnapshotAdmissionInvokesConfiguredLoadersAfterNativeWarmup(string selectedLoader) {
        byte[] data = FontData(false);
        var native = new OfficeFontFaceCollection();
        Assert.True(native.TryAddImmutableSnapshotBounded("Native", data, OfficeFontFaceDescriptor.Regular,
            int.MaxValue, out _, out string? nativeError), nativeError);
        bool providerOnly = selectedLoader == "ProviderOnly";
        var provider = new DecliningProvider();
        int resolverCalls = 0;
        var configured = new OfficeFontFaceCollection();
        if (providerOnly) configured.FontProgramProvider = provider;
        else configured.FontVariationResolver = _ => { resolverCalls++; return null; };
        Assert.True(configured.TryAddImmutableSnapshotBounded("Configured", data, OfficeFontFaceDescriptor.Regular,
            int.MaxValue, out int decodedBytes, out string? error), error);
        Assert.Equal(data.Length * 2, decodedBytes);
        Assert.Equal(providerOnly ? 1 : 0, provider.Calls);
        Assert.Equal(providerOnly ? 0 : 1, resolverCalls);
        Assert.NotSame(Assert.Single(native.Faces).Program, Assert.Single(configured.Faces).Program);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImmutableSnapshotTemplatesDoNotExtendSourceOrProgramLifetime(bool cff) {
        (WeakReference Source, WeakReference Program) references = CreateCollectibleSnapshotTemplate(cff);
        GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking: true, compacting: true);
        GC.WaitForPendingFinalizers();
        GC.Collect(GC.MaxGeneration, GCCollectionMode.Forced, blocking: true, compacting: true);

        Assert.False(references.Source.IsAlive);
        Assert.False(references.Program.IsAlive);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ImmutableSnapshotAdmissionKeepsTrackedAndVariableProgramsScoped(bool variable) {
        byte[] data = variable
            ? File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "RobotoFlex.ttf"))
            : ManagedTextShapingTestAssets.CreateTrackingFont(ManagedTextShapingTestAssets.CreateTrackingTable(
                new[] { 10D, 20D }, new[] { -1D, 1D },
                new[] { new short[] { -120, -40 }, new short[] { -80, 40 } }));
        var first = new OfficeFontFaceCollection();
        var second = new OfficeFontFaceCollection();
        Assert.True(first.TryAddImmutableSnapshotBounded("First", data, OfficeFontFaceDescriptor.Regular,
            int.MaxValue, out int firstBytes, out string? firstError), firstError);
        Assert.True(second.TryAddImmutableSnapshotBounded("Second", data, OfficeFontFaceDescriptor.Regular,
            int.MaxValue, out int secondBytes, out string? secondError), secondError);
        Assert.Equal(data.Length * 2, firstBytes);
        Assert.Equal(firstBytes, secondBytes);
        Assert.NotSame(Assert.Single(first.Faces).Program, Assert.Single(second.Faces).Program);
    }

    private static void AssertEquivalentFontUse(OfficeFontFace reference, OfficeFontFace actual, string text) {
        Assert.Equal(reference.FamilyName, actual.FamilyName);
        Assert.Equal(reference.ResourceFamilyName, actual.ResourceFamilyName);
        Assert.Equal(reference.Style, actual.Style);
        Assert.Equal(reference.Descriptor, actual.Descriptor);
        Assert.Equal(reference.ContainerFormat, actual.ContainerFormat);
        Assert.Equal(reference.CanEmbedAsStaticPdfFont, actual.CanEmbedAsStaticPdfFont);
        Assert.Equal(reference.Data, actual.Data);
        Assert.Equal(reference.Program.Fingerprint, actual.Program.Fingerprint);
        Assert.Equal(reference.Program.DisplayName, actual.Program.DisplayName);
        Assert.Equal(reference.Program.CollectionIndex, actual.Program.CollectionIndex);
        Assert.Equal(reference.Program.UnitsPerEm, actual.Program.UnitsPerEm);
        Assert.Equal(reference.Program.IsOpenTypeCff, actual.Program.IsOpenTypeCff);
        Assert.Equal(reference.Program.ProvidesComplexTextLayout, actual.Program.ProvidesComplexTextLayout);
        Assert.Equal(reference.Program.LineSpacingRatio, actual.Program.LineSpacingRatio);
        Assert.Equal(reference.Program.LineHeight(12D), actual.Program.LineHeight(12D));
        Assert.Equal(reference.Program.Measure(text, 12D), actual.Program.Measure(text, 12D), 9);
        List<List<OfficePoint>> expectedContours = reference.Program.GetTextContours(text, 3.25D, 7.5D, 12D);
        List<List<OfficePoint>> actualContours = actual.Program.GetTextContours(text, 3.25D, 7.5D, 12D);
        Assert.NotEmpty(expectedContours);
        Assert.Equal(expectedContours.Count, actualContours.Count);
        for (int index = 0; index < expectedContours.Count; index++) {
            Assert.Equal<OfficePoint>(expectedContours[index], actualContours[index]);
        }
    }

    [MethodImpl(MethodImplOptions.NoInlining)]
    private static (WeakReference Source, WeakReference Program) CreateCollectibleSnapshotTemplate(bool cff) {
        byte[] snapshot = (byte[])FontData(cff).Clone();
        var first = new OfficeFontFaceCollection();
        Assert.True(first.TryAddImmutableSnapshotBounded("Lifetime", snapshot, OfficeFontFaceDescriptor.Regular,
            int.MaxValue, out int firstBytes, out string? firstError), firstError);
        var second = new OfficeFontFaceCollection();
        Assert.True(second.TryAddImmutableSnapshotBounded("Lifetime Warm", snapshot, OfficeFontFaceDescriptor.Regular,
            int.MaxValue, out int secondBytes, out string? secondError), secondError);
        Assert.Equal(snapshot.Length * (cff ? 3 : 2), firstBytes);
        Assert.Equal(firstBytes, secondBytes);
        OfficeFontFace firstFace = Assert.Single(first.Faces);
        Assert.Same(firstFace.Program, Assert.Single(second.Faces).Program);
        return (new WeakReference(snapshot), new WeakReference(firstFace.Program));
    }

    private static byte[] FontData(bool cff) => cff
        ? File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf"))
        : ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs('A', 'B');

    private sealed class DecliningProvider : IOfficeFontProgramProvider {
        internal int Calls { get; private set; }
        public OfficeFontProgramLoadResult? TryLoad(OfficeFontProgramLoadRequest request) {
            Calls++;
            return null;
        }
    }
}
