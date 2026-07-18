namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Represents one entry in a binary PowerPoint document sound collection.</summary>
    public sealed class LegacyPptSound {
        internal LegacyPptSound(uint id, string name, string? extension,
            int? builtInId, byte[] dataBytes) {
            if (id == 0) throw new ArgumentOutOfRangeException(nameof(id));
            Id = id;
            Name = name ?? throw new ArgumentNullException(nameof(name));
            Extension = extension;
            BuiltInId = builtInId;
            DataBytes = dataBytes ?? throw new ArgumentNullException(nameof(dataBytes));
        }

        /// <summary>Gets the positive document-level identifier referenced by slide-show records.</summary>
        public uint Id { get; }

        /// <summary>Gets the display name stored with the sound.</summary>
        public string Name { get; }

        /// <summary>Gets the four-character binary file extension, when present.</summary>
        public string? Extension { get; }

        /// <summary>Gets the optional built-in PowerPoint sound identifier (100 through 125).</summary>
        public int? BuiltInId { get; }

        /// <summary>Gets the exact embedded WAV or AIFF payload.</summary>
        public byte[] DataBytes { get; }

        /// <summary>Gets whether this entry contains embedded audio bytes.</summary>
        public bool HasData => DataBytes.Length > 0;

        /// <summary>Gets the media content type inferred from the binary extension or payload signature.</summary>
        public string? ContentType {
            get {
                string normalized = (Extension ?? string.Empty).TrimEnd('\0')
                    .TrimStart('.').ToLowerInvariant();
                if (normalized == "wav" || normalized == "wave") return "audio/wav";
                if (normalized == "aif" || normalized == "aiff") return "audio/aiff";
                if (DataBytes.Length >= 12
                    && DataBytes[0] == (byte)'R' && DataBytes[1] == (byte)'I'
                    && DataBytes[2] == (byte)'F' && DataBytes[3] == (byte)'F'
                    && DataBytes[8] == (byte)'W' && DataBytes[9] == (byte)'A'
                    && DataBytes[10] == (byte)'V' && DataBytes[11] == (byte)'E') {
                    return "audio/wav";
                }
                if (DataBytes.Length >= 12
                    && DataBytes[0] == (byte)'F' && DataBytes[1] == (byte)'O'
                    && DataBytes[2] == (byte)'R' && DataBytes[3] == (byte)'M'
                    && ((DataBytes[8] == (byte)'A' && DataBytes[9] == (byte)'I'
                         && DataBytes[10] == (byte)'F' && DataBytes[11] == (byte)'F')
                        || (DataBytes[8] == (byte)'A' && DataBytes[9] == (byte)'I'
                            && DataBytes[10] == (byte)'F' && DataBytes[11] == (byte)'C'))) {
                    return "audio/aiff";
                }
                return null;
            }
        }
    }
}
