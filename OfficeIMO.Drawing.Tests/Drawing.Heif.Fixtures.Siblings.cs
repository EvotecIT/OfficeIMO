using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
namespace OfficeIMO.Tests;
internal static partial class CoreHeifFixtures {
    // Storage 0 is a single absolute mdat extent; 1 is idat-relative; 2 declares no location; 3 splits mdat into two extents.
    internal static byte[] CreateHeifMetadataSiblings(byte[] exif, string xmp, int exifStorage, int xmpStorage, string encoding = "", ushort exifProtection = 0, ushort xmpProtection = 0, byte[]? xmpBytes = null) {
        byte[][] payloads = { Combine(UInt32BigEndian(6), Encoding.ASCII.GetBytes("Exif\0\0"), exif), xmpBytes ?? Encoding.UTF8.GetBytes(xmp) };
        int[] storage = { exifStorage, xmpStorage };
        byte[] ftyp = Box("ftyp", Combine(Encoding.ASCII.GetBytes("heic"), UInt32BigEndian(0), Encoding.ASCII.GetBytes("mif1heic")));
        byte[] idatPayload = Combine(payloads.Where((_, index) => storage[index] == 1).ToArray());
        byte[] mdatPayload = Combine(payloads.Where((_, index) => storage[index] == 0 || storage[index] == 3).ToArray());
        byte[] BuildMeta(uint dataOffset) {
            byte[] exifInfo = FullBox("infe", 2, Combine(UInt16BigEndian(1), UInt16BigEndian(exifProtection), Encoding.ASCII.GetBytes("Exif"), new byte[] { 0 }));
            byte[] xmpInfo = FullBox("infe", 2, Combine(UInt16BigEndian(2), UInt16BigEndian(xmpProtection), Encoding.ASCII.GetBytes("mime"), new byte[] { 0 }, Encoding.ASCII.GetBytes("application/rdf+xml"), new byte[] { 0 }, Encoding.ASCII.GetBytes(encoding), new byte[] { 0 }));
            var locations = new List<byte[]>();
            uint mdatOffset = dataOffset, idatOffset = 0;
            for (int index = 0; index < payloads.Length; index++) {
                if (storage[index] == 2) { continue; }
                bool usesMdat = storage[index] == 0 || storage[index] == 3;
                uint offset = usesMdat ? mdatOffset : idatOffset;
                uint length = (uint)payloads[index].Length;
                byte[] extents = storage[index] == 3
                    ? Combine(UInt32BigEndian(offset), UInt32BigEndian(length / 2), UInt32BigEndian(offset + length / 2), UInt32BigEndian(length - length / 2))
                    : Combine(UInt32BigEndian(offset), UInt32BigEndian(length));
                locations.Add(Combine(UInt16BigEndian((ushort)(index + 1)), UInt16BigEndian((ushort)(usesMdat ? 0 : 1)), UInt16BigEndian(0), UInt16BigEndian((ushort)(storage[index] == 3 ? 2 : 1)), extents));
                if (usesMdat) { mdatOffset += length; } else { idatOffset += length; }
            }
            byte[] iinf = FullBox("iinf", 0, Combine(UInt16BigEndian(2), exifInfo, xmpInfo));
            byte[] iloc = FullBox("iloc", 1, Combine(new byte[] { 0x44, 0x00 }, UInt16BigEndian((ushort)locations.Count), Combine(locations.ToArray())));
            byte[] idat = idatPayload.Length == 0 ? Array.Empty<byte>() : Box("idat", idatPayload);
            return FullBox("meta", 0, Combine(iinf, iloc, idat));
        }
        byte[] placeholder = BuildMeta(0);
        byte[] meta = BuildMeta((uint)(ftyp.Length + placeholder.Length + 8));
        return Combine(ftyp, meta, Box("mdat", mdatPayload));
    }
}
