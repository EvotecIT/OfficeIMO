using OfficeIMO.Core.Internal;
using System.Text;

namespace OfficeIMO.Project;

/// <summary>Constructs native table schemas and empty storage from format definitions, without a source file.</summary>
internal static class ProjectNativeCreation {
    internal static OfficeCompoundFile Empty(DateTime start, long budget, CancellationToken token, ProjectNativeProfile profile) {
        var streams = new Dictionary<string, byte[]>(); var properties = new ProjectNativePropertySet(profile == ProjectNativeProfile.Mpp8);
        void Integer(uint id, int value) => properties.Set(id, 4, BitConverter.GetBytes(value));
        void Short(uint id, short value) => properties.Set(id, 2, BitConverter.GetBytes(value));
        void Text(uint id, string value) => properties.Set(id, 0, Encoding.Unicode.GetBytes(value + "\0"));
        if (profile.HasExtendedRecords) Integer(0x35400016, profile.Version);
        Text(0x024003e8, "TBkndTask,TBkndRsc,TBkndCal,TBkndAssn,TBkndCons" + (profile == ProjectNativeProfile.Mpp8 ? "" : ",TBkndOutlCode"));
        Text(0x02400008, "OfficeIMO project"); Text(0x0240000e, "Standard");
        Integer(0x02400002, unchecked((int)Date(start))); Integer(0x02400003, unchecked((int)Date(start)));
        Short(0x02400004, 1); Integer(0x0240001d, 480); Integer(0x0240001e, 2400); Short(0x0240138f, 20);
        Text(0x024013bb, "USD"); Text(0x02400010, "$"); Short(0x02400012, 2);
        Short(0x0240001c, 4800); Short(0x02400021, 10200);
        if (profile == ProjectNativeProfile.Mpp8) {
            uint ordinal = 0;
            foreach (var table in ProjectNativeSchema8Catalog.Tables()) {
                token.ThrowIfCancellationRequested(); ordinal++;
                string prefix = profile.DataRoot + "/TBknd" + table.Name + "/";
                streams.Add(prefix + "FixFix   0", Array.Empty<byte>());
                if (table.Name != "Cons") streams.Add(prefix + "FixDeferFix   0", BitConverter.GetBytes(-1));
                Integer(0x01000000 | ordinal, 0); properties.Set(0x02000000 | ordinal, 0x10000, table.Schema.Serialize(budget, token));
            }
        }
        var schemas = profile == ProjectNativeProfile.Mpp8 ? Array.Empty<(string, ProjectNativeSchema)>() : profile.Schemas();
        for (int index = 0; index < schemas.Length; index++) {
            token.ThrowIfCancellationRequested(); var item = schemas[index]; uint table = index == 5 ? 8u : (uint)index + 1;
            string prefix = profile.DataRoot + "/TBknd" + item.Item1 + "/";
            streams.Add(prefix + "FixedMeta", Header(16, profile.HasExtendedRecords)); streams.Add(prefix + "FixedData", Array.Empty<byte>());
            if (profile.HasExtendedRecords) { streams.Add(prefix + "Fixed2Meta", Header(16, true)); streams.Add(prefix + "Fixed2Data", Array.Empty<byte>()); }
            streams.Add(prefix + "VarMeta", Header(24, false)); streams.Add(prefix + "Var2Data", Array.Empty<byte>());
            Integer(0x01000000 | table, 0); Integer(0x02000000 | table, 0); Integer(0x00800000 | table, 0);
            Integer(0x00010000 | table, 0); Integer(0x04000000 | table, item.Item2.PrimaryCount); Integer(0x00400000 | table, item.Item2.Width);
            if (profile.HasExtendedRecords) {
                Integer(0x00030000 | table, item.Item2.Count); Integer(0x00040000 | table, item.Item2.SecondaryWidth); Integer(0x00050000 | table, 1);
            }
            properties.Set(0x03000014 + (uint)index, 0, item.Item2.Bytes(false)); if (profile.HasExtendedRecords) properties.Set(0x00020014 + (uint)index, 0, item.Item2.Bytes(true));
        }
        var presentation = new ProjectNativePropertySet(); presentation.Set(0x024003e8, 0, Encoding.Unicode.GetBytes("\0"));
        streams.Add(profile.PresentationRoot + "/Props", presentation.Serialize(budget, token)); streams.Add(profile.Properties, properties.Serialize(budget, token));
        var header = new ProjectNativePropertySet();
        string producer = profile.HasExtendedRecords ? "16,0,0,0\0" : profile == ProjectNativeProfile.Mpp8 ? "8,0,0,0\0" : "9,0,0,0\0";
        header.Set(0x35400010, 0, Encoding.Unicode.GetBytes(producer)); header.Set(0x3540000c, 0, Encoding.Unicode.GetBytes(producer));
        header.Set(0x35400008, 0, Encoding.Unicode.GetBytes("Project1\0")); header.Set(0x35400000, 0, new byte[1]); header.Set(0x35400001, 0, new byte[1]);
        header.Set(0x35400002, 2, BitConverter.GetBytes((short)1)); header.Set(0x3540000f, 2, new byte[2]);
        streams.Add(profile.Header, header.Serialize(budget, token));
        streams.Add("\u0001CompObj", OfficeOleCompoundObjectWriter.Write(new Guid("74b78f3a-c8c8-11d1-be11-00c04fb6faf1"), "Microsoft.Project 16.0", "MSProject.MPP" + profile.Version, "MSProject.Project.9"));
        streams.Add(OfficeOlePropertySetWriter.SummaryInformationStreamName, OfficeOlePropertySetWriter.CreatePropertySet(
            (OfficeOlePropertySetWriter.SummaryInformationFormatId, OfficeOlePropertySetWriter.CreateSection(new[] {
                OfficeOleProperty.Integer(1, (short)1200), OfficeOleProperty.String(2, "OfficeIMO project"), OfficeOleProperty.String(18, "OfficeIMO.Project") }))));
        return new OfficeCompoundFile(streams, streams.Select(pair => new OfficeCompoundFileEntry(pair.Key.Split('/').Last(), pair.Key, 2, pair.Value.Length)).ToArray(),
            new OfficeCompoundFileEntry("Root Entry", "Root Entry", 5, 0, classId: new Guid("74b78f3a-c8c8-11d1-be11-00c04fb6faf1")));
    }
    internal static uint Date(DateTime date) {
        if (date.Kind != DateTimeKind.Unspecified) throw new ArgumentException("Native dates require an unspecified local project clock.", nameof(date));
        if (date.Ticks % (TimeSpan.TicksPerMinute / 10) != 0) throw new ArgumentException("Native date precision is one tenth of a minute.", nameof(date));
        int days = (date.Date - new DateTime(1983, 12, 31)).Days;
        if (days < 0 || days > ushort.MaxValue) throw new ArgumentOutOfRangeException(nameof(date));
        uint value = (uint)days << 16 | (uint)(date.TimeOfDay.Ticks / (TimeSpan.TicksPerMinute / 10));
        if (value == 0) throw new ArgumentOutOfRangeException(nameof(date), "The native zero date is reserved for an absent value.");
        return value;
    }
    private static byte[] Header(int size, bool extended) {
        var bytes = new byte[size]; Buffer.BlockCopy(BitConverter.GetBytes(0xfadfadbau), 0, bytes, 0, 4);
        if (extended) bytes[4] = 4; return bytes;
    }
}
