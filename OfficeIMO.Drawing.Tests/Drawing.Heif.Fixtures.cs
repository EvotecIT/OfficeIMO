using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

using OfficeIMO.Drawing;
using ExifTag = OfficeIMO.Drawing.OfficeExifTag;
using Xunit;
namespace OfficeIMO.Tests;
internal static partial class CoreHeifFixtures {
    internal static byte[] CreateMinimalHeifWithExif(byte[] tiffPayload) {
        byte[] exifItemData = Combine(
            UInt32BigEndian(6),
            Encoding.ASCII.GetBytes("Exif\0\0"),
            tiffPayload);

        byte[] ftyp = Box("ftyp", Combine(
            Encoding.ASCII.GetBytes("heic"),
            UInt32BigEndian(0),
            Encoding.ASCII.GetBytes("heic"),
            Encoding.ASCII.GetBytes("mif1")));

        byte[] metaWithPlaceholder = CreateMetaBox(0, exifItemData.Length);
        uint exifOffset = (uint)(ftyp.Length + metaWithPlaceholder.Length + 8);
        byte[] meta = CreateMetaBox(exifOffset, exifItemData.Length);
        byte[] mdat = Box("mdat", exifItemData);

        return Combine(ftyp, meta, mdat);
    }

    internal static byte[] CreateMinimalHeifWithPrimaryImageExifAndXmp(uint width, uint height, byte[] tiffPayload, string xmp) {
        byte[] exifItemData = Combine(
            UInt32BigEndian(6),
            Encoding.ASCII.GetBytes("Exif\0\0"),
            tiffPayload);
        byte[] xmpItemData = Encoding.UTF8.GetBytes(xmp);

        byte[] ftyp = Box("ftyp", Combine(
            Encoding.ASCII.GetBytes("heic"),
            UInt32BigEndian(0),
            Encoding.ASCII.GetBytes("heic"),
            Encoding.ASCII.GetBytes("mif1")));

        byte[] metaWithPlaceholder = CreateMetaBoxWithPrimaryImageExifAndXmp(0, exifItemData.Length, 0, xmpItemData.Length, width, height);
        uint exifOffset = (uint)(ftyp.Length + metaWithPlaceholder.Length + 8);
        uint xmpOffset = (uint)(exifOffset + exifItemData.Length);
        byte[] meta = CreateMetaBoxWithPrimaryImageExifAndXmp(exifOffset, exifItemData.Length, xmpOffset, xmpItemData.Length, width, height);
        byte[] mdat = Box("mdat", Combine(exifItemData, xmpItemData));

        return Combine(ftyp, meta, mdat);
    }

    internal static byte[] CreateMinimalHeifWithPrimaryImageTransformProperties(uint width, uint height, byte[] tiffPayload) {
        byte[] exifItemData = Combine(
            UInt32BigEndian(6),
            Encoding.ASCII.GetBytes("Exif\0\0"),
            tiffPayload);

        byte[] ftyp = Box("ftyp", Combine(
            Encoding.ASCII.GetBytes("heic"),
            UInt32BigEndian(0),
            Encoding.ASCII.GetBytes("heic"),
            Encoding.ASCII.GetBytes("mif1")));

        byte[] metaWithPlaceholder = CreateMetaBoxWithPrimaryImageTransformProperties(0, exifItemData.Length, width, height);
        uint exifOffset = (uint)(ftyp.Length + metaWithPlaceholder.Length + 8);
        byte[] meta = CreateMetaBoxWithPrimaryImageTransformProperties(exifOffset, exifItemData.Length, width, height);
        byte[] mdat = Box("mdat", exifItemData);

        return Combine(ftyp, meta, mdat);
    }

    internal static byte[] CreateMinimalHeifWithPrimaryImageAndExif(uint width, uint height, byte[] tiffPayload) {
        byte[] exifItemData = Combine(
            UInt32BigEndian(6),
            Encoding.ASCII.GetBytes("Exif\0\0"),
            tiffPayload);

        byte[] ftyp = Box("ftyp", Combine(
            Encoding.ASCII.GetBytes("heic"),
            UInt32BigEndian(0),
            Encoding.ASCII.GetBytes("heic"),
            Encoding.ASCII.GetBytes("mif1")));

        byte[] metaWithPlaceholder = CreateMetaBoxWithPrimaryImage(0, exifItemData.Length, width, height);
        uint exifOffset = (uint)(ftyp.Length + metaWithPlaceholder.Length + 8);
        byte[] meta = CreateMetaBoxWithPrimaryImage(exifOffset, exifItemData.Length, width, height);
        byte[] mdat = Box("mdat", exifItemData);

        return Combine(ftyp, meta, mdat);
    }

    internal static byte[] CreateMinimalHeifWithAuxiliaryImage() {
        byte[] ftyp = Box("ftyp", Combine(
            Encoding.ASCII.GetBytes("heic"),
            UInt32BigEndian(0),
            Encoding.ASCII.GetBytes("heic"),
            Encoding.ASCII.GetBytes("mif1")));
        byte[] meta = CreateMetaBoxWithAuxiliaryImage();

        return Combine(ftyp, meta);
    }

    internal static byte[] CreateMinimalHeifWithIdatXmp(string xmp) {
        byte[] xmpItemData = Encoding.UTF8.GetBytes(xmp);
        byte[] ftyp = Box("ftyp", Combine(
            Encoding.ASCII.GetBytes("heic"),
            UInt32BigEndian(0),
            Encoding.ASCII.GetBytes("heic"),
            Encoding.ASCII.GetBytes("mif1")));
        byte[] meta = CreateMetaBoxWithIdatXmp(xmpItemData);

        return Combine(ftyp, meta);
    }

    internal static byte[] CreateMinimalHeifWithUnlocatedXmp() {
        byte[] ftyp = Box("ftyp", Combine(
            Encoding.ASCII.GetBytes("heic"),
            UInt32BigEndian(0),
            Encoding.ASCII.GetBytes("heic"),
            Encoding.ASCII.GetBytes("mif1")));
        byte[] meta = CreateMetaBoxWithUnlocatedXmp();

        return Combine(ftyp, meta);
    }

    internal static byte[] CreateMinimalHeifWithIdatExif(byte[] tiffPayload) {
        byte[] exifItemData = Combine(
            UInt32BigEndian(6),
            Encoding.ASCII.GetBytes("Exif\0\0"),
            tiffPayload);
        byte[] ftyp = Box("ftyp", Combine(
            Encoding.ASCII.GetBytes("heic"),
            UInt32BigEndian(0),
            Encoding.ASCII.GetBytes("heic"),
            Encoding.ASCII.GetBytes("mif1")));
        byte[] meta = CreateMetaBoxWithIdatExif(exifItemData);

        return Combine(ftyp, meta);
    }

    internal static byte[] CreateExifPayload(string software) {
        var profile = new OfficeImageMetadata();
        profile.SetExifValue(ExifTag.Software, software);
        return Assert.IsType<byte[]>(profile.EncodeExifProfile());
    }

    internal static byte[] CreateMinimalHeifWithoutExif() {
        byte[] ftyp = Box("ftyp", Combine(
            Encoding.ASCII.GetBytes("heic"),
            UInt32BigEndian(0),
            Encoding.ASCII.GetBytes("heic"),
            Encoding.ASCII.GetBytes("mif1")));

        byte[] iinf = FullBox("iinf", 0, UInt16BigEndian(0));
        byte[] meta = FullBox("meta", 0, iinf);
        return Combine(ftyp, meta);
    }

    internal static byte[] CreateMetaBox(uint exifOffset, int exifLength) {
        byte[] infe = FullBox("infe", 2, Combine(
            UInt16BigEndian(1),
            UInt16BigEndian(0),
            Encoding.ASCII.GetBytes("Exif"),
            new byte[] { 0 }));

        byte[] iinf = FullBox("iinf", 0, Combine(
            UInt16BigEndian(1),
            infe));

        byte[] iloc = FullBox("iloc", 1, Combine(
            new byte[] { 0x44, 0x00 },
            UInt16BigEndian(1),
            UInt16BigEndian(1),
            UInt16BigEndian(0),
            UInt16BigEndian(0),
            UInt16BigEndian(1),
            UInt32BigEndian(exifOffset),
            UInt32BigEndian((uint)exifLength)));

        return FullBox("meta", 0, Combine(iinf, iloc));
    }

    internal static byte[] CreateMetaBoxWithPrimaryImage(uint exifOffset, int exifLength, uint width, uint height) {
        byte[] imageInfe = FullBox("infe", 2, Combine(
            UInt16BigEndian(1),
            UInt16BigEndian(0),
            Encoding.ASCII.GetBytes("hvc1"),
            new byte[] { 0 }));
        byte[] exifInfe = FullBox("infe", 2, Combine(
            UInt16BigEndian(2),
            UInt16BigEndian(0),
            Encoding.ASCII.GetBytes("Exif"),
            new byte[] { 0 }));

        byte[] iinf = FullBox("iinf", 0, Combine(
            UInt16BigEndian(2),
            imageInfe,
            exifInfe));

        byte[] pitm = FullBox("pitm", 0, UInt16BigEndian(1));
        byte[] iloc = FullBox("iloc", 1, Combine(
            new byte[] { 0x44, 0x00 },
            UInt16BigEndian(1),
            UInt16BigEndian(2),
            UInt16BigEndian(0),
            UInt16BigEndian(0),
            UInt16BigEndian(1),
            UInt32BigEndian(exifOffset),
            UInt32BigEndian((uint)exifLength)));
        byte[] ispe = FullBox("ispe", 0, Combine(UInt32BigEndian(width), UInt32BigEndian(height)));
        byte[] ipco = Box("ipco", ispe);
        byte[] ipma = FullBox("ipma", 0, Combine(
            UInt32BigEndian(1),
            UInt16BigEndian(1),
            new byte[] { 1, 1 }));
        byte[] iprp = Box("iprp", Combine(ipco, ipma));

        return FullBox("meta", 0, Combine(pitm, iinf, iloc, iprp));
    }

    internal static byte[] CreateMetaBoxWithPrimaryImageTransformProperties(uint exifOffset, int exifLength, uint width, uint height) {
        byte[] imageInfe = FullBox("infe", 2, Combine(
            UInt16BigEndian(1),
            UInt16BigEndian(0),
            Encoding.ASCII.GetBytes("hvc1"),
            new byte[] { 0 }));
        byte[] exifInfe = FullBox("infe", 2, Combine(
            UInt16BigEndian(2),
            UInt16BigEndian(0),
            Encoding.ASCII.GetBytes("Exif"),
            new byte[] { 0 }));

        byte[] iinf = FullBox("iinf", 0, Combine(
            UInt16BigEndian(2),
            imageInfe,
            exifInfe));

        byte[] pitm = FullBox("pitm", 0, UInt16BigEndian(1));
        byte[] iloc = FullBox("iloc", 1, Combine(
            new byte[] { 0x44, 0x00 },
            UInt16BigEndian(1),
            UInt16BigEndian(2),
            UInt16BigEndian(0),
            UInt16BigEndian(0),
            UInt16BigEndian(1),
            UInt32BigEndian(exifOffset),
            UInt32BigEndian((uint)exifLength)));
        byte[] ispe = FullBox("ispe", 0, Combine(UInt32BigEndian(width), UInt32BigEndian(height)));
        byte[] irot = Box("irot", new byte[] { 1 });
        byte[] imir = Box("imir", new byte[] { 0 });
        byte[] pasp = Box("pasp", Combine(UInt32BigEndian(4), UInt32BigEndian(3)));
        byte[] pixi = FullBox("pixi", 0, new byte[] { 3, 8, 8, 8 });
        byte[] colr = Box("colr", Combine(
            Encoding.ASCII.GetBytes("nclx"),
            UInt16BigEndian(1),
            UInt16BigEndian(13),
            UInt16BigEndian(6),
            new byte[] { 0x80 }));
        byte[] hvcC = Box("hvcC", new byte[] { 1, 2, 3, 4 });
        byte[] ipco = Box("ipco", Combine(ispe, irot, imir, pasp, pixi, colr, hvcC));
        byte[] ipma = FullBox("ipma", 0, Combine(
            UInt32BigEndian(1),
            UInt16BigEndian(1),
            new byte[] { 7, 0x81, 2, 3, 4, 5, 6, 7 }));
        byte[] iprp = Box("iprp", Combine(ipco, ipma));

        return FullBox("meta", 0, Combine(pitm, iinf, iloc, iprp));
    }

    internal static byte[] CreateMetaBoxWithAuxiliaryImage() {
        byte[] imageInfe = FullBox("infe", 2, Combine(
            UInt16BigEndian(1),
            UInt16BigEndian(0),
            Encoding.ASCII.GetBytes("hvc1"),
            new byte[] { 0 }));
        byte[] auxiliaryInfe = FullBox("infe", 2, Combine(
            UInt16BigEndian(2),
            UInt16BigEndian(0),
            Encoding.ASCII.GetBytes("hvc1"),
            new byte[] { 0 }));

        byte[] iinf = FullBox("iinf", 0, Combine(
            UInt16BigEndian(2),
            imageInfe,
            auxiliaryInfe));
        byte[] pitm = FullBox("pitm", 0, UInt16BigEndian(1));
        byte[] auxReference = Box("auxl", Combine(
            UInt16BigEndian(2),
            UInt16BigEndian(1),
            UInt16BigEndian(1)));
        byte[] iref = FullBox("iref", 0, auxReference);
        byte[] ispe = FullBox("ispe", 0, Combine(UInt32BigEndian(320), UInt32BigEndian(240)));
        byte[] auxC = FullBox("auxC", 0, Combine(
            Encoding.ASCII.GetBytes("urn:mpeg:hevc:2015:auxid:1"),
            new byte[] { 0, 0x10, 0x20 }));
        byte[] ipco = Box("ipco", Combine(ispe, auxC));
        byte[] ipma = FullBox("ipma", 0, Combine(
            UInt32BigEndian(2),
            UInt16BigEndian(1),
            new byte[] { 1, 1 },
            UInt16BigEndian(2),
            new byte[] { 1, 2 }));
        byte[] iprp = Box("iprp", Combine(ipco, ipma));

        return FullBox("meta", 0, Combine(pitm, iinf, iref, iprp));
    }

    internal static byte[] CreateMetaBoxWithIdatXmp(byte[] xmpItemData) {
        byte[] xmpInfe = FullBoxWithFlags("infe", 2, 1, Combine(
            UInt16BigEndian(1),
            UInt16BigEndian(0),
            Encoding.ASCII.GetBytes("mime"),
            new byte[] { 0 },
            Encoding.ASCII.GetBytes("application/rdf+xml"),
            new byte[] { 0 },
            Array.Empty<byte>(),
            new byte[] { 0 }));
        byte[] iinf = FullBox("iinf", 0, Combine(
            UInt16BigEndian(1),
            xmpInfe));
        byte[] iloc = FullBox("iloc", 1, Combine(
            new byte[] { 0x44, 0x00 },
            UInt16BigEndian(1),
            UInt16BigEndian(1),
            UInt16BigEndian(1),
            UInt16BigEndian(0),
            UInt16BigEndian(1),
            UInt32BigEndian(0),
            UInt32BigEndian((uint)xmpItemData.Length)));
        byte[] idat = Box("idat", xmpItemData);

        return FullBox("meta", 0, Combine(iinf, iloc, idat));
    }

    internal static byte[] CreateMetaBoxWithUnlocatedXmp() {
        byte[] xmpInfe = FullBoxWithFlags("infe", 2, 1, Combine(
            UInt16BigEndian(1),
            UInt16BigEndian(0),
            Encoding.ASCII.GetBytes("mime"),
            new byte[] { 0 },
            Encoding.ASCII.GetBytes("application/rdf+xml"),
            new byte[] { 0 },
            Array.Empty<byte>(),
            new byte[] { 0 }));
        byte[] iinf = FullBox("iinf", 0, Combine(
            UInt16BigEndian(1),
            xmpInfe));

        return FullBox("meta", 0, iinf);
    }

    internal static byte[] CreateMetaBoxWithIdatExif(byte[] exifItemData) {
        byte[] exifInfe = FullBox("infe", 2, Combine(
            UInt16BigEndian(1),
            UInt16BigEndian(0),
            Encoding.ASCII.GetBytes("Exif"),
            new byte[] { 0 }));
        byte[] iinf = FullBox("iinf", 0, Combine(
            UInt16BigEndian(1),
            exifInfe));
        byte[] iloc = FullBox("iloc", 1, Combine(
            new byte[] { 0x44, 0x00 },
            UInt16BigEndian(1),
            UInt16BigEndian(1),
            UInt16BigEndian(1),
            UInt16BigEndian(0),
            UInt16BigEndian(1),
            UInt32BigEndian(0),
            UInt32BigEndian((uint)exifItemData.Length)));
        byte[] idat = Box("idat", exifItemData);

        return FullBox("meta", 0, Combine(iinf, iloc, idat));
    }

    internal static byte[] CreateMetaBoxWithPrimaryImageExifAndXmp(uint exifOffset, int exifLength, uint xmpOffset, int xmpLength, uint width, uint height) {
        byte[] imageInfe = FullBox("infe", 2, Combine(
            UInt16BigEndian(1),
            UInt16BigEndian(0),
            Encoding.ASCII.GetBytes("hvc1"),
            new byte[] { 0 }));
        byte[] exifInfe = FullBox("infe", 2, Combine(
            UInt16BigEndian(2),
            UInt16BigEndian(0),
            Encoding.ASCII.GetBytes("Exif"),
            new byte[] { 0 }));
        byte[] xmpInfe = FullBoxWithFlags("infe", 2, 1, Combine(
            UInt16BigEndian(3),
            UInt16BigEndian(0),
            Encoding.ASCII.GetBytes("mime"),
            new byte[] { 0 },
            Encoding.ASCII.GetBytes("application/rdf+xml"),
            new byte[] { 0 },
            Array.Empty<byte>(),
            new byte[] { 0 }));

        byte[] iinf = FullBox("iinf", 0, Combine(
            UInt16BigEndian(3),
            imageInfe,
            exifInfe,
            xmpInfe));

        byte[] pitm = FullBox("pitm", 0, UInt16BigEndian(1));
        byte[] exifLocation = Combine(
            UInt16BigEndian(2),
            UInt16BigEndian(0),
            UInt16BigEndian(0),
            UInt16BigEndian(1),
            UInt32BigEndian(exifOffset),
            UInt32BigEndian((uint)exifLength));
        byte[] xmpLocation = Combine(
            UInt16BigEndian(3),
            UInt16BigEndian(0),
            UInt16BigEndian(0),
            UInt16BigEndian(1),
            UInt32BigEndian(xmpOffset),
            UInt32BigEndian((uint)xmpLength));
        byte[] iloc = FullBox("iloc", 1, Combine(
            new byte[] { 0x44, 0x00 },
            UInt16BigEndian(2),
            exifLocation,
            xmpLocation));
        byte[] exifReference = Box("cdsc", Combine(
            UInt16BigEndian(2),
            UInt16BigEndian(1),
            UInt16BigEndian(1)));
        byte[] xmpReference = Box("cdsc", Combine(
            UInt16BigEndian(3),
            UInt16BigEndian(1),
            UInt16BigEndian(1)));
        byte[] iref = FullBox("iref", 0, Combine(exifReference, xmpReference));
        byte[] ispe = FullBox("ispe", 0, Combine(UInt32BigEndian(width), UInt32BigEndian(height)));
        byte[] ipco = Box("ipco", ispe);
        byte[] ipma = FullBox("ipma", 0, Combine(
            UInt32BigEndian(1),
            UInt16BigEndian(1),
            new byte[] { 1, 1 }));
        byte[] iprp = Box("iprp", Combine(ipco, ipma));

        return FullBox("meta", 0, Combine(pitm, iinf, iref, iloc, iprp));
    }

    internal static byte[] FullBox(string type, byte version, byte[] payload) =>
        Box(type, Combine(new byte[] { version, 0, 0, 0 }, payload));

    internal static byte[] FullBoxWithFlags(string type, byte version, uint flags, byte[] payload) =>
        Box(type, Combine(new[] { version, (byte)(flags >> 16), (byte)(flags >> 8), (byte)flags }, payload));

    internal static byte[] Box(string type, byte[] payload) =>
        Combine(UInt32BigEndian((uint)(8 + payload.Length)), Encoding.ASCII.GetBytes(type), payload);

    internal static byte[] UInt16BigEndian(ushort value) =>
        new[] {
            (byte)(value >> 8),
            (byte)value
        };

    internal static byte[] UInt32BigEndian(uint value) =>
        new[] {
            (byte)(value >> 24),
            (byte)(value >> 16),
            (byte)(value >> 8),
            (byte)value
        };

    internal static bool ContainsSequence(byte[] data, byte[] sequence) {
        if (sequence.Length == 0) {
            return true;
        }

        for (int dataIndex = 0; dataIndex <= data.Length - sequence.Length; dataIndex++) {
            bool matched = true;
            for (int sequenceIndex = 0; sequenceIndex < sequence.Length; sequenceIndex++) {
                if (data[dataIndex + sequenceIndex] != sequence[sequenceIndex]) {
                    matched = false;
                    break;
                }
            }

            if (matched) {
                return true;
            }
        }

        return false;
    }

    internal static byte[] Combine(params byte[][] arrays) {
        var result = new byte[arrays.Sum(static a => a.Length)];
        int offset = 0;
        foreach (byte[] array in arrays) {
            Buffer.BlockCopy(array, 0, result, offset, array.Length);
            offset += array.Length;
        }

        return result;
    }
}
