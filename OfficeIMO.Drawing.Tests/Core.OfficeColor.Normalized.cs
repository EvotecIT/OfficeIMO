using System;
using System.Collections.Generic;
using OfficeIMO;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class OfficeColorNormalizedTests {
    [Theory]
    [InlineData(0.123456789012345D, 0.5D, 0.9D, 17)]
    [InlineData(double.Epsilon, 0.5D, 1D, 0)]
    [InlineData(0D, 1D, 0.25D, 255)]
    public void NormalizedChannelsRetainTheirInputAndIndependentByteAlpha(double r, double g, double b, byte alpha) {
        OfficeColor color = OfficeColor.FromNormalizedRgb(r, g, b, alpha);

        Assert.Equal((r, g, b), color.ToNormalizedRgb());
        Assert.Equal(alpha, color.A);
        Assert.False(color.IsThemeRole);
        Assert.Equal(OfficeColorRole.None, color.Role);
        Assert.Null(color.ThemeSlot);
    }

    [Fact]
    public void EveryByteChannelHasAnEqualNormalizedConstructionAndOriginalHash() {
        for (int value = 0; value <= byte.MaxValue; value++) {
            byte r = (byte)value;
            byte g = (byte)(255 - value);
            byte b = (byte)((value * 73) % 256);
            byte alpha = (byte)value;
            OfficeColor bytes = new OfficeColor(r, g, b, alpha);
            OfficeColor normalized = OfficeColor.FromNormalizedRgb(r / 255D, g / 255D, b / 255D, alpha);
            int originalHash = (r << 24) | (g << 16) | (b << 8) | alpha;

            Assert.Equal(bytes, normalized);
            Assert.Equal(normalized, bytes);
            Assert.True(bytes == normalized);
            Assert.False(bytes != normalized);
            Assert.Equal(originalHash, bytes.GetHashCode());
            Assert.Equal(originalHash, normalized.GetHashCode());
            Assert.Equal((r / 255D, g / 255D, b / 255D), bytes.ToNormalizedRgb());
            Assert.Equal(bytes.ToNormalizedRgb(), normalized.ToNormalizedRgb());
        }
    }

    [Fact]
    public void AValueAdjacentToAnExactByteFractionDoesNotCollapseToItsByteCounterpart() {
        double fraction = 128 / 255D;
        double next = BitConverter.Int64BitsToDouble(BitConverter.DoubleToInt64Bits(fraction) + 1);
        OfficeColor bytes = OfficeColor.FromRgb(128, 0, 255);
        OfficeColor nearby = OfficeColor.FromNormalizedRgb(next, 0D, 1D);
        var cache = new Dictionary<OfficeColor, string> { [bytes] = "bytes", [nearby] = "nearby" };

        Assert.Equal(bytes.ToHex(), nearby.ToHex());
        Assert.Equal((next, 0D, 1D), nearby.ToNormalizedRgb());
        Assert.NotEqual(bytes, nearby);
        Assert.Equal(2, cache.Count);
        Assert.Equal("nearby", cache[OfficeColor.FromNormalizedRgb(next, 0D, 1D)]);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void NearbyPreciseChannelsRemainDistinctCacheKeysDespiteIdenticalByteProjection(int channel) {
        double next = BitConverter.Int64BitsToDouble(BitConverter.DoubleToInt64Bits(0.5D) + 1);
        OfficeColor first = OfficeColor.FromNormalizedRgb(0.5D, 0.5D, 0.5D);
        OfficeColor repeated = OfficeColor.FromNormalizedRgb(0.5D, 0.5D, 0.5D);
        OfficeColor nearby = OfficeColor.FromNormalizedRgb(
            channel == 0 ? next : 0.5D, channel == 1 ? next : 0.5D, channel == 2 ? next : 0.5D);
        OfficeColor bytes = OfficeColor.FromRgb(128, 128, 128);
        var cache = new Dictionary<OfficeColor, string> {
            [first] = "precise", [nearby] = "nearby", [bytes] = "bytes"
        };

        Assert.Equal(first.ToHex(), nearby.ToHex());
        Assert.Equal(first.ToHex(), bytes.ToHex());
        Assert.NotEqual(first, nearby);
        Assert.NotEqual(first, bytes);
        Assert.NotEqual(nearby, bytes);
        Assert.True(first != nearby);
        Assert.False(first == nearby);
        Assert.Equal(first, repeated);
        Assert.True(((object)first).Equals(repeated));
        Assert.Equal(first.GetHashCode(), repeated.GetHashCode());
        // These adjacent inputs exercise the hash mixer; unequal colors need not have unique hashes in general.
        Assert.NotEqual(first.GetHashCode(), nearby.GetHashCode());
        Assert.Equal("precise", cache[repeated]);
        Assert.Equal("nearby", cache[nearby]);
        cache[OfficeColor.FromNormalizedRgb(128 / 255D, 128 / 255D, 128 / 255D)] = "updated bytes";
        Assert.Equal("updated bytes", cache[bytes]);
        Assert.Equal(3, cache.Count);
    }

    [Fact]
    public void AlphaRemainsPartOfPreciseLiteralIdentity() {
        OfficeColor transparent = OfficeColor.FromNormalizedRgb(0.5D, 0.25D, 0.75D, 0);
        OfficeColor opaque = OfficeColor.FromNormalizedRgb(0.5D, 0.25D, 0.75D);
        var colors = new HashSet<OfficeColor> { transparent, opaque };

        Assert.Equal(transparent.ToNormalizedRgb(), opaque.ToNormalizedRgb());
        Assert.Equal((byte)0, transparent.A);
        Assert.Equal((byte)255, opaque.A);
        Assert.NotEqual(transparent, opaque);
        Assert.Equal(2, colors.Count);
    }

    [Fact]
    public void ByteGettersAndLiteralFormattingAreDeliberatelyLossyProjections() {
        OfficeColor precise = OfficeColor.FromNormalizedRgb(0.125D, 0.5D, 0.9D, 17);
        OfficeColor projected = OfficeColor.FromRgba(32, 128, 230, 17);

        Assert.Equal((byte)32, precise.R);
        Assert.Equal((byte)128, precise.G);
        Assert.Equal((byte)230, precise.B);
        Assert.Equal("2080E611", precise.ToHex());
        Assert.Equal("2080E6", precise.ToRgbHex());
        Assert.Equal("112080E6", precise.ToArgbHex());
        Assert.Equal("#2080e611", precise.ToCssColor());
        Assert.Equal("#2080E611", precise.ToString());
        Assert.Equal(projected, OfficeColor.ParseHex(precise.ToHex()));
        Assert.Equal(projected, OfficeColor.ParseCss(precise.ToCssColor()));
        Assert.NotEqual(precise, projected);
        Assert.Equal((0.125D, 0.5D, 0.9D), precise.ToNormalizedRgb());
    }

    [Fact]
    public void DefaultAndByteConstructionKeepTheirAlphaAndUnits() {
        OfficeColor zero = default;
        OfficeColor black = new OfficeColor(0, 0, 0);
        double negativeZero = BitConverter.Int64BitsToDouble(long.MinValue);

        Assert.Equal(OfficeColor.Transparent, zero);
        Assert.Equal(0, zero.GetHashCode());
        Assert.Equal((byte)0, zero.A);
        Assert.Equal((0D, 0D, 0D), zero.ToNormalizedRgb());
        Assert.Equal(zero, OfficeColor.FromNormalizedRgb(negativeZero, 0D, 0D, 0));
        Assert.Equal(zero.GetHashCode(), OfficeColor.FromNormalizedRgb(negativeZero, 0D, 0D, 0).GetHashCode());
        Assert.Equal(OfficeColor.Black, black);
        Assert.Equal((byte)255, black.A);
        Assert.Equal(255, black.GetHashCode());
        Assert.NotEqual(zero, black);
        Assert.Equal((128 / 255D, 0D, 1D), new OfficeColor(128, 0, 255).ToNormalizedRgb());
    }

    [Fact]
    public void ThemeRolesStillRequireResolutionAndKeepTheirExistingHash() {
        foreach (OfficeColorRole role in Enum.GetValues(typeof(OfficeColorRole))) {
            if (role == OfficeColorRole.None) continue;
            OfficeColor color = OfficeColor.Parse(role.ToString());

            Assert.Equal(unchecked((int)0x9E3779B9 ^ (int)role), color.GetHashCode());
            Assert.Throws<InvalidOperationException>(() => color.ToNormalizedRgb());
        }
    }

    [Fact]
    public void ThemeResolutionAndValueCopiesRetainPreciseLiteralChannels() {
        OfficeColor precise = OfficeColor.FromNormalizedRgb(0.5D, 0.123456789012345D, 0.9D, 64);
        OfficeColor copy = precise;
        OfficeTheme theme = OfficeTheme.Modern.WithPalette(OfficeTheme.Modern.Palette.With(accent1: precise));
        OfficeColor resolved = OfficeColor.Accent1.Resolve(theme);

        Assert.Equal(precise, copy);
        Assert.Equal(precise.ToNormalizedRgb(), copy.ToNormalizedRgb());
        Assert.Equal(precise, precise.Resolve(theme));
        Assert.Equal(precise, resolved);
        Assert.Equal(precise.GetHashCode(), resolved.GetHashCode());
        Assert.Equal(precise.ToNormalizedRgb(), resolved.ToNormalizedRgb());
        Assert.Equal((byte)64, resolved.A);
        Assert.False(resolved.IsThemeRole);
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.PositiveInfinity)]
    [InlineData(double.NegativeInfinity)]
    [InlineData(-double.Epsilon)]
    [InlineData(1.0000000000000002D)]
    public void InvalidNormalizedChannelsAreRejectedWithTheirParameterName(double invalid) {
        Assert.Equal("r", Assert.Throws<ArgumentOutOfRangeException>(
            () => OfficeColor.FromNormalizedRgb(invalid, 0.5D, 0.5D)).ParamName);
        Assert.Equal("g", Assert.Throws<ArgumentOutOfRangeException>(
            () => OfficeColor.FromNormalizedRgb(0.5D, invalid, 0.5D)).ParamName);
        Assert.Equal("b", Assert.Throws<ArgumentOutOfRangeException>(
            () => OfficeColor.FromNormalizedRgb(0.5D, 0.5D, invalid)).ParamName);
    }
}
