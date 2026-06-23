namespace OfficeIMO.Excel.LegacyXls.Biff {
    internal readonly struct BiffRecord {
        internal BiffRecord(ushort type, int offset, byte[] payload) {
            Type = type;
            Offset = offset;
            Payload = payload;
        }

        internal ushort Type { get; }

        internal int Offset { get; }

        internal byte[] Payload { get; }
    }
}
