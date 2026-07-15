namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Identifies transition effects defined by PowerPoint 97-2003.</summary>
    public enum LegacyPptTransitionEffect : byte {
        /// <summary>Cut, optionally through black.</summary>
        Cut = 0,
        /// <summary>Randomly selected transition.</summary>
        Random = 1,
        /// <summary>Horizontal or vertical blinds.</summary>
        Blinds = 2,
        /// <summary>Horizontal or vertical checker pattern.</summary>
        Checker = 3,
        /// <summary>Cover from a direction.</summary>
        Cover = 4,
        /// <summary>Dissolve.</summary>
        Dissolve = 5,
        /// <summary>Fade through black.</summary>
        Fade = 6,
        /// <summary>Uncover toward a direction.</summary>
        Uncover = 7,
        /// <summary>Horizontal or vertical random bars.</summary>
        RandomBars = 8,
        /// <summary>Diagonal strips.</summary>
        Strips = 9,
        /// <summary>Directional wipe.</summary>
        Wipe = 10,
        /// <summary>Box inward or outward.</summary>
        Box = 11,
        /// <summary>Horizontal or vertical split.</summary>
        Split = 13,
        /// <summary>Diamond.</summary>
        Diamond = 17,
        /// <summary>Plus.</summary>
        Plus = 18,
        /// <summary>Wedge.</summary>
        Wedge = 19,
        /// <summary>Directional push.</summary>
        Push = 20,
        /// <summary>Horizontal or vertical comb.</summary>
        Comb = 21,
        /// <summary>Newsflash.</summary>
        Newsflash = 22,
        /// <summary>Smooth alpha fade directly between slides.</summary>
        AlphaFade = 23,
        /// <summary>Wheel with radial divisions.</summary>
        Wheel = 26,
        /// <summary>Circle.</summary>
        Circle = 27,
        /// <summary>Undefined effect marker.</summary>
        Undefined = 255
    }

    /// <summary>Represents a complete SlideShowSlideInfoAtom.</summary>
    public sealed class LegacyPptTransition {
        internal LegacyPptTransition(int slideTimeMilliseconds, uint soundId,
            byte effectDirection, byte effectType, ushort rawFlags, byte speed) {
            SlideTimeMilliseconds = slideTimeMilliseconds;
            SoundId = soundId;
            EffectDirection = effectDirection;
            RawEffectType = effectType;
            Effect = Enum.IsDefined(typeof(LegacyPptTransitionEffect), effectType)
                ? (LegacyPptTransitionEffect)effectType
                : null;
            RawFlags = rawFlags;
            Speed = speed;
        }

        /// <summary>Gets the automatic-advance delay in milliseconds.</summary>
        public int SlideTimeMilliseconds { get; }
        /// <summary>Gets the referenced transition-sound identifier.</summary>
        public uint SoundId { get; }
        /// <summary>Gets the effect-specific direction or variant byte.</summary>
        public byte EffectDirection { get; }
        /// <summary>Gets the raw legacy effect type.</summary>
        public byte RawEffectType { get; }
        /// <summary>Gets the typed legacy effect, or null when undefined.</summary>
        public LegacyPptTransitionEffect? Effect { get; }
        /// <summary>Gets the complete transition flags field.</summary>
        public ushort RawFlags { get; }
        /// <summary>Gets the raw speed value: 0 slow, 1 medium, or 2 fast.</summary>
        public byte Speed { get; }
        /// <summary>Gets whether clicking can advance the slide.</summary>
        public bool ManualAdvance => (RawFlags & 0x0001) != 0;
        /// <summary>Gets whether the slide is hidden during the show.</summary>
        public bool Hidden => (RawFlags & 0x0004) != 0;
        /// <summary>Gets whether the referenced transition sound is played.</summary>
        public bool PlaySound => (RawFlags & 0x0010) != 0;
        /// <summary>Gets whether the transition sound loops.</summary>
        public bool LoopSound => (RawFlags & 0x0040) != 0;
        /// <summary>Gets whether a currently playing sound is stopped.</summary>
        public bool StopSound => (RawFlags & 0x0100) != 0;
        /// <summary>Gets whether the slide automatically advances.</summary>
        public bool AutoAdvance => (RawFlags & 0x0400) != 0;
        /// <summary>Gets whether the cursor is visible during the slide show.</summary>
        public bool CursorVisible => (RawFlags & 0x1000) != 0;
    }
}
