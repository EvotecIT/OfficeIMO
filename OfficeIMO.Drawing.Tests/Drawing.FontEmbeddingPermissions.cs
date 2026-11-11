using System.Buffers.Binary;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests {
    public sealed class DrawingFontEmbeddingPermissionsTests {
        [Theory]
        [InlineData(0x0000, (int)OfficeFontEmbeddingLevel.Installable, false, false, true)]
        [InlineData(0x0002, (int)OfficeFontEmbeddingLevel.Restricted, false, false, false)]
        [InlineData(0x0004, (int)OfficeFontEmbeddingLevel.PreviewAndPrint, false, false, true)]
        [InlineData(0x0008, (int)OfficeFontEmbeddingLevel.Editable, false, false, true)]
        [InlineData(0x0006, (int)OfficeFontEmbeddingLevel.PreviewAndPrint, false, false, true)]
        [InlineData(0x000A, (int)OfficeFontEmbeddingLevel.Editable, false, false, true)]
        [InlineData(0x000E, (int)OfficeFontEmbeddingLevel.Editable, false, false, true)]
        [InlineData(0x0100, (int)OfficeFontEmbeddingLevel.Installable, true, false, true)]
        [InlineData(0x0200, (int)OfficeFontEmbeddingLevel.Installable, false, true, false)]
        [InlineData(0x0302, (int)OfficeFontEmbeddingLevel.Restricted, true, true, false)]
        public void FsTypeResolvesToTheLeastRestrictivePermission(int fsType, int level, bool noSubsetting, bool bitmapOnly, bool outlines) {
            byte[] font = WithFsType(Font(), (ushort)fsType);

            OfficeFontEmbeddingPermissions permissions = OfficeFontEmbeddingPermissions.TryRead(font)!.Value;

            Assert.True(permissions.IsDeclared);
            Assert.Equal((OfficeFontEmbeddingLevel)level, permissions.Level);
            Assert.Equal(noSubsetting, permissions.NoSubsetting);
            Assert.Equal(bitmapOnly, permissions.BitmapOnly);
            Assert.Equal(outlines, permissions.AllowsOutlineEmbedding);
            Assert.Equal(outlines, permissions.OutlineEmbeddingRefusal is null);
        }

        [Fact]
        public void MissingOs2DeclaresNoRestrictionAndCollectionsAreReadPerFace() {
            byte[] withoutOs2 = Font();
            RenameTable(withoutOs2, "OS/2", "OS/3");
            OfficeFontEmbeddingPermissions undeclared = OfficeFontEmbeddingPermissions.TryRead(withoutOs2)!.Value;
            Assert.False(undeclared.IsDeclared);
            Assert.True(undeclared.AllowsOutlineEmbedding);

            // OS/2 versions 0 and 1 define only bits 0-3, so a bitmap-only bit there is ignored.
            byte[] version1 = WithFsType(Font(), 0x0200);
            BinaryPrimitives.WriteUInt16BigEndian(version1.AsSpan(FindTable(version1, "OS/2")), 1);
            Assert.True(OfficeFontEmbeddingPermissions.TryRead(version1)!.Value.AllowsOutlineEmbedding);

            byte[] collection = Collection(Font(), WithFsType(Font(), 0x0002));
            Assert.Equal(OfficeFontEmbeddingLevel.Installable, OfficeFontEmbeddingPermissions.TryRead(collection, 0)!.Value.Level);
            Assert.Equal(OfficeFontEmbeddingLevel.Restricted, OfficeFontEmbeddingPermissions.TryRead(collection, 1)!.Value.Level);
            Assert.Equal(OfficeFontEmbeddingLevel.Restricted, OfficeFontEmbeddingPermissions.TryReadFile(collection)!.Value.Level);
            Assert.Null(OfficeFontEmbeddingPermissions.TryReadFile(Encoding.ASCII.GetBytes("not a font program")));
        }

        [Theory]
        [InlineData(0x0000, true)]
        [InlineData(0x0006, true)]
        [InlineData(0x0002, false)]
        [InlineData(0x0200, false)]
        public void SvgExportEmbedsOnlyFacesWhoseLicenceAllowsIt(int fsType, bool embedded) {
            var drawing = new OfficeDrawing(120D, 30D);
            drawing.Fonts.Add("Licensed Face", WithFsType(Font(), (ushort)fsType));
            drawing.AddText("\u6C38\u5B57\u56FD", 0D, 0D, 120D, 30D, new OfficeFontInfo("Licensed Face", 12D));

            string svg = OfficeDrawingSvgExporter.ToSvg(drawing);

            Assert.Equal(embedded, svg.Contains("@font-face", StringComparison.Ordinal));
            Assert.Contains("Licensed Face", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ProviderKeptWoffFacesAreReadForTheirLicence() {
            byte[] restricted = WithFsType(Font(), 0x0002);
            var fonts = new OfficeFontFaceCollection { FontProgramProvider = new SfntProvider(restricted) };

            Assert.True(fonts.TryAdd("Licensed Woff", ManagedTextShapingTestAssets.CreateWoff(restricted)));

            Assert.Equal(OfficeFontContainerFormat.Woff, fonts.Faces[0].ContainerFormat);
            Assert.False(fonts.Faces[0].AllowsOutlineEmbedding);

            // A face whose licence cannot be read at all is not embedded either.
            var unreadable = new OfficeFontFaceCollection { FontProgramProvider = new SfntProvider(Font()) };
            Assert.True(unreadable.TryAdd("Unreadable", Encoding.ASCII.GetBytes("provider-specific font program bytes")));
            Assert.False(unreadable.Faces[0].AllowsOutlineEmbedding);
        }

        // Supplies a program but no static sfnt data, so the face keeps the WOFF container bytes.
        private sealed class SfntProvider : IOfficeFontProgramProvider {
            private readonly byte[] _sfnt;
            internal SfntProvider(byte[] sfnt) => _sfnt = sfnt;
            public OfficeFontProgramLoadResult? TryLoad(OfficeFontProgramLoadRequest request) =>
                new OfficeFontProgramLoadResult(OfficeTrueTypeFont.TryLoad(_sfnt)!, _sfnt.Length);
        }

        internal static byte[] Font() =>
            File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "NotoSansSC-BaselineSubset.ttf"));

        internal static byte[] WithFsType(byte[] font, ushort fsType) {
            BinaryPrimitives.WriteUInt16BigEndian(font.AsSpan(FindTable(font, "OS/2") + 8), fsType);
            return font;
        }

        private static int FindTable(byte[] font, string tag) {
            int count = BinaryPrimitives.ReadUInt16BigEndian(font.AsSpan(4));
            for (int index = 0; index < count; index++) {
                int record = 12 + index * 16;
                if (Encoding.ASCII.GetString(font, record, 4) == tag) return (int)BinaryPrimitives.ReadUInt32BigEndian(font.AsSpan(record + 8));
            }
            throw new InvalidOperationException("Missing " + tag + " table.");
        }

        private static void RenameTable(byte[] font, string tag, string replacement) {
            int count = BinaryPrimitives.ReadUInt16BigEndian(font.AsSpan(4));
            for (int index = 0; index < count; index++) {
                int record = 12 + index * 16;
                if (Encoding.ASCII.GetString(font, record, 4) == tag) Encoding.ASCII.GetBytes(replacement, 0, 4, font, record);
            }
        }

        // Places each whole font after a ttcf header, moving its table offsets by the face's position.
        private static byte[] Collection(params byte[][] faces) {
            int headerLength = 12 + faces.Length * 4;
            var output = new List<byte>(new byte[headerLength]);
            byte[] header = new byte[headerLength];
            Encoding.ASCII.GetBytes("ttcf", 0, 4, header, 0);
            BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(4), 0x00010000);
            BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(8), (uint)faces.Length);
            for (int face = 0; face < faces.Length; face++) {
                while (output.Count % 4 != 0) output.Add(0);
                int start = output.Count;
                BinaryPrimitives.WriteUInt32BigEndian(header.AsSpan(12 + face * 4), (uint)start);
                byte[] copy = (byte[])faces[face].Clone();
                int count = BinaryPrimitives.ReadUInt16BigEndian(copy.AsSpan(4));
                for (int index = 0; index < count; index++) {
                    Span<byte> offset = copy.AsSpan(12 + index * 16 + 8, 4);
                    BinaryPrimitives.WriteUInt32BigEndian(offset, BinaryPrimitives.ReadUInt32BigEndian(offset) + (uint)start);
                }
                output.AddRange(copy);
            }
            byte[] result = output.ToArray();
            Buffer.BlockCopy(header, 0, result, 0, headerLength);
            return result;
        }
    }
}
