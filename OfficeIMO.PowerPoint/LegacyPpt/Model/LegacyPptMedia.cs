namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Identifies a native binary PowerPoint audio or movie object.</summary>
    public enum LegacyPptMediaKind {
        /// <summary>An external AVI movie.</summary>
        AviMovie,
        /// <summary>An external Media Control Interface movie.</summary>
        MciMovie,
        /// <summary>An external MIDI audio file.</summary>
        MidiAudio,
        /// <summary>A compact-disc audio range.</summary>
        CdAudio,
        /// <summary>WAV audio embedded through the document sound collection.</summary>
        EmbeddedWaveAudio,
        /// <summary>An external WAV audio file.</summary>
        LinkedWaveAudio
    }

    /// <summary>
    /// Represents one external-media definition from a binary PowerPoint file.
    /// </summary>
    public sealed class LegacyPptMedia {
        internal LegacyPptMedia(uint id, LegacyPptMediaKind kind,
            bool loop, bool rewind, bool narration, string? path,
            uint? soundId, int? durationMilliseconds,
            uint? cdStart, uint? cdEnd, LegacyPptSound? sound) {
            if (id == 0) throw new ArgumentOutOfRangeException(nameof(id));
            Id = id;
            Kind = kind;
            Loop = loop;
            Rewind = rewind;
            Narration = narration;
            Path = path;
            SoundId = soundId;
            DurationMilliseconds = durationMilliseconds;
            CdStart = cdStart;
            CdEnd = cdEnd;
            Sound = sound;
        }

        /// <summary>Gets the document-wide external-object identifier.</summary>
        public uint Id { get; }
        /// <summary>Gets the native media representation.</summary>
        public LegacyPptMediaKind Kind { get; }
        /// <summary>Gets whether playback repeats continuously.</summary>
        public bool Loop { get; }
        /// <summary>Gets whether the media rewinds after playback.</summary>
        public bool Rewind { get; }
        /// <summary>Gets whether the audio is recorded narration.</summary>
        public bool Narration { get; }
        /// <summary>Gets the external UNC or local path, when present.</summary>
        public string? Path { get; }
        /// <summary>Gets the referenced document sound identifier.</summary>
        public uint? SoundId { get; }
        /// <summary>Gets the embedded WAV playback duration, in milliseconds.</summary>
        public int? DurationMilliseconds { get; }
        /// <summary>Gets the raw packed CD start time.</summary>
        public uint? CdStart { get; }
        /// <summary>Gets the raw packed CD end time.</summary>
        public uint? CdEnd { get; }
        /// <summary>Gets the referenced embedded sound entry, when resolved.</summary>
        public LegacyPptSound? Sound { get; }
        /// <summary>Gets whether this object can project as editable embedded WAV media.</summary>
        public bool HasProjectableAudio => Kind ==
                LegacyPptMediaKind.EmbeddedWaveAudio
            && Sound?.HasData == true
            && string.Equals(Sound.ContentType, "audio/wav",
                StringComparison.OrdinalIgnoreCase);

        /// <summary>Returns a defensive copy of the embedded audio bytes.</summary>
        public byte[] GetData() => Sound?.DataBytes == null
            ? Array.Empty<byte>()
            : (byte[])Sound.DataBytes.Clone();
    }
}
