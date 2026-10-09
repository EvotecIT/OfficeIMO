using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace OfficeIMO.Tests;

internal static partial class CoreHeifFixtures {
    // Associations: 0=primary/thumbnail, 1=ambiguous primary, 2=none,
    // 3=reversed direction, 4=only another image, 5=no primary declaration.
    internal static byte[] CreateHeifAssociationGraph(bool largeIds, int associations = 0,
        bool uniqueMetadata = false, int malformedReference = 0, int malformedCollection = 0,
        bool protectedPrimaryExif = false, bool encodedPrimaryXmp = false,
        int propertyContainerMode = 0, int propertyAssociationMode = 0) {
        uint primary = largeIds ? 70000U : 1U;
        uint thumbnail = primary + 1, exifPrimary = primary + 2, exifThumbnail = primary + 3;
        uint xmpPrimary = primary + 4, xmpThumbnail = primary + 5;
        byte[] Id(uint id) => largeIds ? UInt32BigEndian(id) : UInt16BigEndian((ushort)id);
        var payloads = new Dictionary<uint, byte[]> {
            [primary] = Encoding.ASCII.GetBytes("primary-encoded-pixels"),
            [thumbnail] = Encoding.ASCII.GetBytes("thumbnail-encoded-pixels"),
            [exifPrimary] = Combine(UInt32BigEndian(6), Encoding.ASCII.GetBytes("Exif\0\0"), CreateExifPayload("Primary")),
            [xmpPrimary] = Encoding.UTF8.GetBytes("primary-XMP")
        };
        if (!uniqueMetadata) {
            payloads[exifThumbnail] = Combine(UInt32BigEndian(6), Encoding.ASCII.GetBytes("Exif\0\0"), CreateExifPayload("Thumbnail"));
            payloads[xmpThumbnail] = Encoding.UTF8.GetBytes("thumbnail-XMP");
        }
        // Thumbnail metadata deliberately precedes primary metadata in both iinf and iloc.
        uint[] ids = uniqueMetadata
            ? new[] { primary, thumbnail, exifPrimary, xmpPrimary }
            : new[] { primary, thumbnail, exifThumbnail, exifPrimary, xmpThumbnail, xmpPrimary };
        byte[] Infe(uint id) {
            string type = id == primary || id == thumbnail ? "hvc1" : id == exifPrimary || id == exifThumbnail ? "Exif" : "mime";
            ushort protection = id == exifPrimary && protectedPrimaryExif ? (ushort)1 : (ushort)0;
            byte[] mime = type == "mime"
                ? Combine(Encoding.ASCII.GetBytes("application/rdf+xml"), new byte[] { 0 },
                    Encoding.ASCII.GetBytes(id == xmpPrimary && encodedPrimaryXmp ? "gzip" : ""), new byte[] { 0 })
                : Array.Empty<byte>();
            byte[] body = Combine(Id(id), UInt16BigEndian(protection), Encoding.ASCII.GetBytes(type), new byte[] { 0 }, mime);
            if (malformedCollection == 1 && id == ids[ids.Length - 1]) {
                body = new byte[largeIds ? 3 : 1]; // Truncated mandatory item ID.
            }
            return FullBox("infe", largeIds ? (byte)3 : (byte)2, body);
        }
        byte[] Reference(uint from, uint to, bool broken) {
            byte[] body = Combine(Id(from), UInt16BigEndian(1), Id(to));
            if (broken && malformedReference == 1) {
                body = new byte[largeIds ? 3 : 1];
            } else if (broken && malformedReference == 2) {
                body = Combine(Id(from), new byte[] { 0 });
            } else if (broken && malformedReference == 3) {
                body = Combine(Id(from), UInt16BigEndian(2), Id(to));
            }
            return Box("cdsc", body);
        }
        byte[] BuildMeta(uint offset) {
            byte[] iinf = FullBox("iinf", 0, Combine(UInt16BigEndian((ushort)ids.Length), Combine(ids.Select(Infe).ToArray())));
            var entries = new List<byte[]>();
            uint at = offset;
            foreach (uint id in ids) {
                byte[] entry = Combine(Id(id), largeIds ? UInt16BigEndian(0) : Array.Empty<byte>(),
                    UInt16BigEndian(0), UInt16BigEndian(1), UInt32BigEndian(at), UInt32BigEndian((uint)payloads[id].Length));
                if (malformedCollection == 2 && id == ids[ids.Length - 1]) {
                    entry = entry.Take(entry.Length - 2).ToArray();
                }
                entries.Add(entry);
                at += (uint)payloads[id].Length;
            }
            byte[] iloc = FullBox("iloc", largeIds ? (byte)2 : (byte)0,
                Combine(new byte[] { 0x44, 0 }, largeIds ? UInt32BigEndian((uint)ids.Length) : UInt16BigEndian((ushort)ids.Length), Combine(entries.ToArray())));
            var refs = new List<byte[]>();
            if (associations != 2) {
                foreach (uint id in ids.Where(value => value != primary && value != thumbnail)) {
                    bool primaryMetadata = id == exifPrimary || id == xmpPrimary;
                    uint target = associations == 1 || primaryMetadata ? primary : thumbnail;
                    if (associations == 4) {
                        target = thumbnail;
                    }
                    refs.Add(associations == 3
                        ? Reference(target, id, id == exifPrimary)
                        : Reference(id, target, id == exifPrimary));
                }
            }
            byte[] references = Combine(refs.ToArray());
            if (malformedReference == 4) {
                references = Combine(references, new byte[] { 1, 2, 3 });
            }
            byte[] iref = associations == 2 ? Array.Empty<byte>() : FullBox("iref", largeIds ? (byte)1 : (byte)0, references);
            byte[] pitm = associations == 5 ? Array.Empty<byte>() : FullBox("pitm", largeIds ? (byte)1 : (byte)0, Id(primary));
            byte[] propertyEntries = Combine(Id(primary), new byte[] { 1, 1 });
            uint propertyEntryCount = 1;
            if (malformedCollection == 3) {
                propertyEntries = Combine(propertyEntries, Id(thumbnail), new byte[] { 2, 1 });
                propertyEntryCount = 2;
            }
            byte[] propertyContainer = propertyContainerMode == 2
                ? Array.Empty<byte>()
                : Box(propertyContainerMode == 1 ? "free" : "ipco", Combine(
                    FullBox("ispe", 0, Combine(UInt32BigEndian(640), UInt32BigEndian(480))),
                    propertyAssociationMode == 2 ? Box("irot", new byte[] { 1 }) : Array.Empty<byte>()));
            byte[] propertyAssociations = FullBox("ipma", largeIds ? (byte)1 : (byte)0,
                Combine(UInt32BigEndian(propertyEntryCount), propertyEntries));
            if (propertyAssociationMode != 0) {
                uint secondItem = propertyAssociationMode == 2 ? primary : thumbnail;
                byte associationCount = propertyAssociationMode == 3 ? (byte)2 : (byte)1;
                byte propertyIndex = propertyAssociationMode == 2 ? (byte)2 : (byte)1;
                propertyAssociations = Combine(propertyAssociations,
                    FullBox("ipma", largeIds ? (byte)1 : (byte)0,
                        Combine(UInt32BigEndian(1), Id(secondItem), new[] { associationCount, propertyIndex })));
            }
            byte[] iprp = Box("iprp", Combine(propertyContainer, propertyAssociations));
            return FullBox("meta", 0, Combine(pitm, iinf, iref, iloc, iprp));
        }
        byte[] ftyp = Box("ftyp", Combine(Encoding.ASCII.GetBytes("heic"), UInt32BigEndian(0), Encoding.ASCII.GetBytes("mif1heic")));
        byte[] placeholder = BuildMeta(0);
        return Combine(ftyp, BuildMeta((uint)(ftyp.Length + placeholder.Length + 8)), Box("mdat", Combine(ids.Select(id => payloads[id]).ToArray())));
    }
}
