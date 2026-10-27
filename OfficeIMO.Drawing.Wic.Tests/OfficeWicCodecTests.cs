using System;
using System.Buffers.Binary;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Xunit;

namespace OfficeIMO.Drawing.Wic.Tests;

[Collection("WIC platform registration")]
public sealed class OfficeWicCodecTests {
    private static byte[] Probe() => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Probe.heic"));

    [WindowsFact]
    public void Availability_ReportsRegisteredDecoderAndRealPixelDecode() {
        OfficeRasterCodecAvailability result = new OfficeWicCodec().GetAvailability();
        Assert.Equal(result.Reason == OfficeRasterCodecUnavailableReason.None, result.IsAvailable);
        Assert.False(string.IsNullOrWhiteSpace(result.Message) && !result.IsAvailable);
        string? expected = Environment.GetEnvironmentVariable("OFFICEIMO_WIC_EXPECT_AVAILABILITY");
        if (!string.IsNullOrEmpty(expected)) {
            Assert.True(Enum.TryParse(expected, out OfficeRasterCodecUnavailableReason reason),
                "OFFICEIMO_WIC_EXPECT_AVAILABILITY must name an OfficeRasterCodecUnavailableReason value.");
            Assert.True(result.Reason == reason, $"Expected {reason}; got {result.Reason}: {result.Message}");
        }
        if (result.IsAvailable) {
            Assert.True(new OfficeWicCodec().TryDecode(Probe(), "image/heic", out OfficeRasterImage? image));
            Assert.NotNull(image);
        }
    }

    [WindowsFact]
    public void Register_ReplacesPlatformCodec() {
        IOfficeRasterImageCodec? previous = OfficeRasterImageCodecs.Platform;
        try {
            OfficeWicCodec.Register();
            Assert.IsType<OfficeWicCodec>(OfficeRasterImageCodecs.Platform);
        } finally {
            OfficeRasterImageCodecs.Platform = previous;
        }
    }

    [WindowsFact]
    public void TryDecode_RejectsInvalidUnrelatedAndTruncatedInput() {
        var codec = new OfficeWicCodec();
        byte[][] inputs = { Array.Empty<byte>(), new byte[24], new byte[] { 137, 80, 78, 71 }, Probe()[..32], null! };
        foreach (byte[] input in inputs) {
            Assert.False(codec.TryDecode(input, "image/heic", out OfficeRasterImage? image));
            Assert.Null(image);
        }
        // An attacker-controlled extended ftyp length must be checked before any COM invocation.
        byte[] oversizedBox = Probe();
        BinaryPrimitives.WriteUInt32BigEndian(oversizedBox, 1);
        BinaryPrimitives.WriteUInt64BigEndian(oversizedBox.AsSpan(8), ulong.MaxValue);
        Assert.False(codec.TryDecode(oversizedBox, null, out _));
    }

    [WindowsFact]
    public void TryDecode_RejectsInputOver128MiB() {
        byte[] bytes = new byte[128 * 1024 * 1024 + 1];
        Probe().CopyTo(bytes, 0);
        Assert.False(new OfficeWicCodec().TryDecode(bytes, "image/heic", out OfficeRasterImage? image));
        Assert.Null(image);
    }

    [WindowsHeicFact]
    public void TryDecode_ProducesStraightRgbaFromSystemEncodedHeic() {
        var codec = new OfficeWicCodec();
        // Content type is a hint; magic bytes own format selection.
        Assert.True(codec.TryDecode(Probe(), "application/octet-stream", out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.Equal(16, image.Width);
        Assert.Equal(16, image.Height);
        byte[] pixels = image.GetPixels();
        Assert.Equal(16 * 16 * 4, pixels.Length);
        for (int i = 0; i < pixels.Length; i += 4) {
            Assert.InRange((int)pixels[i], 86, 106);
            Assert.InRange((int)pixels[i + 1], 134, 154);
            Assert.InRange((int)pixels[i + 2], 182, 202);
            Assert.Equal(255, pixels[i + 3]);
        }
    }

    [WindowsHeicFact]
    public void TryDecode_UsesBoundedNativeDimensionsWhenContainerSizeIsCorrupt() {
        byte[] bytes = Probe();
        // ispe is a FullBox: type + four version/flag bytes + width + height.
        int ispe = bytes.AsSpan().IndexOf("ispe"u8);
        Assert.True(ispe >= 0);
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(ispe + 8), 50_000_001);
        BinaryPrimitives.WriteUInt32BigEndian(bytes.AsSpan(ispe + 12), 1);
        bool decoded = new OfficeWicCodec().TryDecode(bytes, null, out OfficeRasterImage? image);
        if (decoded) {
            // A system decoder may ignore contradictory ispe metadata in favor of the bitstream.
            // That is safe only when the returned native dimensions still describe the tiny image.
            Assert.NotNull(image);
            Assert.Equal(16, image.Width);
            Assert.Equal(16, image.Height);
        } else {
            Assert.Null(image);
        }
    }

    [WindowsHeicFact]
    public void TryDecode_IsReentrantAcrossMtaThreads() {
        var codec = new OfficeWicCodec();
        byte[] bytes = Probe();
        Parallel.For(0, 12, _ => {
            Assert.True(codec.TryDecode(bytes, null, out OfficeRasterImage? image));
            Assert.NotNull(image);
            Assert.Equal(16, image.Width);
        });
    }

    [WindowsHeicFact]
    public void TryDecode_PreservesExistingStaApartmentAndReleasesNativeReferences() {
        if (!OperatingSystem.IsWindows()) return;
        Exception? failure = null;
        var thread = new Thread(() => {
            try {
                var codec = new OfficeWicCodec();
                for (int i = 0; i < 8; i++) {
                    byte[] bytes = Probe();
                    Assert.True(codec.TryDecode(bytes, null, out _));
                    GC.Collect(); // Subsequent calls must not depend on a released pinned input or an old RCW.
                }
                Assert.Equal(ApartmentState.STA, Thread.CurrentThread.GetApartmentState());
            } catch (Exception ex) {
                failure = ex;
            }
        }) { IsBackground = true };
        thread.SetApartmentState(ApartmentState.STA);
        thread.Start();
        Assert.True(thread.Join(TimeSpan.FromSeconds(30)), "The WIC STA decode did not complete within 30 seconds.");
        Assert.Null(failure);
    }
}
