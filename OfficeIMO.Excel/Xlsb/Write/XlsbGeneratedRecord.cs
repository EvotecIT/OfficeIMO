namespace OfficeIMO.Excel.Xlsb.Write {
    /// <summary>Represents one generated BIFF12 record before framing.</summary>
    internal sealed class XlsbGeneratedRecord {
        internal XlsbGeneratedRecord(int type, byte[] payload) {
            Type = type;
            Payload = payload ?? throw new ArgumentNullException(nameof(payload));
        }

        internal int Type { get; }

        internal byte[] Payload { get; }
    }
}
