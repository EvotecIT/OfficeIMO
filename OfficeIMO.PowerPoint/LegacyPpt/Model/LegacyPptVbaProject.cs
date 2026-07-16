namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents the complete VBA compound storage referenced by a binary presentation.</summary>
    public sealed class LegacyPptVbaProject {
        private readonly byte[] _projectBytes;

        internal LegacyPptVbaProject(uint persistId, bool wasCompressed,
            byte[] projectBytes) {
            PersistId = persistId;
            WasCompressed = wasCompressed;
            _projectBytes = (byte[])projectBytes.Clone();
        }

        /// <summary>Gets the persist identifier referenced by the document VBA information atom.</summary>
        public uint PersistId { get; }

        /// <summary>Gets whether the source persist object used the compressed storage form.</summary>
        public bool WasCompressed { get; }

        /// <summary>Gets the length of the decompressed VBA compound storage.</summary>
        public int Length => _projectBytes.Length;

        /// <summary>Returns a copy of the complete <c>vbaProject.bin</c> compound storage.</summary>
        public byte[] GetBytes() => (byte[])_projectBytes.Clone();
    }
}
