namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Specifies how a legacy shape or its child content is built.</summary>
    public enum LegacyPptAnimationBuildType : byte {
        /// <summary>The shape has no animation build.</summary>
        None = 0x00,
        /// <summary>The shape animates as one object.</summary>
        AsOneObject = 0x01,
        /// <summary>Paragraphs through level 1 animate separately.</summary>
        ByLevel1Paragraph = 0x02,
        /// <summary>Paragraphs through level 2 animate separately.</summary>
        ByLevel2Paragraph = 0x03,
        /// <summary>Paragraphs through level 3 animate separately.</summary>
        ByLevel3Paragraph = 0x04,
        /// <summary>Paragraphs through level 4 animate separately.</summary>
        ByLevel4Paragraph = 0x05,
        /// <summary>Paragraphs through level 5 animate separately.</summary>
        ByLevel5Paragraph = 0x06,
        /// <summary>Chart series animate separately.</summary>
        ChartBySeries = 0x07,
        /// <summary>Chart categories animate separately.</summary>
        ChartByCategory = 0x08,
        /// <summary>Chart elements animate in series order.</summary>
        ChartByElementInSeries = 0x09,
        /// <summary>Chart elements animate in category order.</summary>
        ChartByElementInCategory = 0x0A,
        /// <summary>A placeholder follows its master animation build.</summary>
        FollowMaster = 0xFE
    }

    /// <summary>Specifies a classic PowerPoint 97-2003 animation effect.</summary>
    public enum LegacyPptAnimationEffect : byte {
        /// <summary>Cut or appear.</summary>
        Cut = 0x00,
        /// <summary>A random entrance effect.</summary>
        Random = 0x01,
        /// <summary>Blinds.</summary>
        Blinds = 0x02,
        /// <summary>Checkerboard.</summary>
        Checker = 0x03,
        /// <summary>Cover.</summary>
        Cover = 0x04,
        /// <summary>Dissolve.</summary>
        Dissolve = 0x05,
        /// <summary>Fade.</summary>
        Fade = 0x06,
        /// <summary>Pull or uncover.</summary>
        Pull = 0x07,
        /// <summary>Random bars.</summary>
        RandomBars = 0x08,
        /// <summary>Strips.</summary>
        Strips = 0x09,
        /// <summary>Wipe.</summary>
        Wipe = 0x0A,
        /// <summary>Box zoom.</summary>
        Zoom = 0x0B,
        /// <summary>Fly, crawl, stretch, zoom, swivel, or spiral variant.</summary>
        Fly = 0x0C,
        /// <summary>Split.</summary>
        Split = 0x0D,
        /// <summary>Flash.</summary>
        Flash = 0x0E,
        /// <summary>Diamond.</summary>
        Diamond = 0x11,
        /// <summary>Plus.</summary>
        Plus = 0x12,
        /// <summary>Wedge.</summary>
        Wedge = 0x13,
        /// <summary>Wheel.</summary>
        Wheel = 0x1A,
        /// <summary>Circle.</summary>
        Circle = 0x1B
    }

    /// <summary>Specifies the state applied after a legacy animation completes.</summary>
    public enum LegacyPptAnimationAfterEffect : byte {
        /// <summary>No after-effect.</summary>
        None = 0x00,
        /// <summary>Dim to the authored color.</summary>
        Dim = 0x01,
        /// <summary>Hide on the next mouse click.</summary>
        HideOnNextClick = 0x02,
        /// <summary>Hide immediately.</summary>
        HideImmediately = 0x03
    }

    /// <summary>Specifies how text is subdivided within a legacy animation build.</summary>
    public enum LegacyPptTextBuildSubEffect : byte {
        /// <summary>Animate the selected text at once.</summary>
        AllAtOnce = 0x00,
        /// <summary>Animate word by word.</summary>
        ByWord = 0x01,
        /// <summary>Animate character by character.</summary>
        ByCharacter = 0x02
    }

    /// <summary>Represents the complete classic animation atom attached to a binary shape.</summary>
    public sealed class LegacyPptAnimation {
        internal LegacyPptAnimation(uint rawDimColor, uint rawFlags, uint soundIdReference,
            int delayMilliseconds, short order, ushort slideCount,
            LegacyPptAnimationBuildType buildType, LegacyPptAnimationEffect effect,
            byte effectDirection, LegacyPptAnimationAfterEffect afterEffect,
            LegacyPptTextBuildSubEffect textBuildSubEffect, byte oleVerb,
            ushort rawUnused, bool hasSoundOverride) {
            RawDimColor = rawDimColor;
            RawFlags = rawFlags;
            SoundIdReference = soundIdReference;
            DelayMilliseconds = delayMilliseconds;
            Order = order;
            SlideCount = slideCount;
            BuildType = buildType;
            Effect = effect;
            EffectDirection = effectDirection;
            AfterEffect = afterEffect;
            TextBuildSubEffect = textBuildSubEffect;
            OleVerb = oleVerb;
            RawUnused = rawUnused;
            HasSoundOverride = hasSoundOverride;
        }

        /// <summary>Gets the raw ColorIndexStruct used by a dim after-effect.</summary>
        public uint RawDimColor { get; }

        /// <summary>Gets the complete authored flag word, including preserved reserved bits.</summary>
        public uint RawFlags { get; }

        /// <summary>Gets whether the effect plays in reverse.</summary>
        public bool Reverse => ReadFlag(0);

        /// <summary>Gets whether the effect starts automatically rather than on click.</summary>
        public bool Automatic => ReadFlag(2);

        /// <summary>Gets whether an associated sound is played.</summary>
        public bool PlaysSound => ReadFlag(4);

        /// <summary>Gets whether all playing sounds are stopped when the effect begins.</summary>
        public bool StopsSound => ReadFlag(6);

        /// <summary>Gets whether the associated sound, media, or OLE verb plays on shape click.</summary>
        public bool PlaysOnShapeClick => ReadFlag(8);

        /// <summary>Gets whether media or OLE playback blocks other slide-show actions.</summary>
        public bool Synchronous => ReadFlag(10);

        /// <summary>Gets whether a media or OLE shape is hidden while it is not playing.</summary>
        public bool HiddenWhileNotPlaying => ReadFlag(12);

        /// <summary>Gets whether the shape background participates in the effect.</summary>
        public bool AnimateBackground => ReadFlag(14);

        /// <summary>Gets the document sound identifier referenced by the animation.</summary>
        public uint SoundIdReference { get; }

        /// <summary>Gets the automatic-start delay in milliseconds.</summary>
        public int DelayMilliseconds { get; }

        /// <summary>Gets the authored animation order; -2 means follow the master placeholder.</summary>
        public short Order { get; }

        /// <summary>Gets the number-of-slides field used by media animations.</summary>
        public ushort SlideCount { get; }

        /// <summary>Gets the shape, text, or chart build mode.</summary>
        public LegacyPptAnimationBuildType BuildType { get; }

        /// <summary>Gets the classic visual effect.</summary>
        public LegacyPptAnimationEffect Effect { get; }

        /// <summary>Gets the effect-specific direction or variant.</summary>
        public byte EffectDirection { get; }

        /// <summary>Gets the state applied after the effect completes.</summary>
        public LegacyPptAnimationAfterEffect AfterEffect { get; }

        /// <summary>Gets whether text is animated all at once, by word, or by character.</summary>
        public LegacyPptTextBuildSubEffect TextBuildSubEffect { get; }

        /// <summary>Gets the authored OLE verb byte.</summary>
        public byte OleVerb { get; }

        /// <summary>Gets the raw trailing two bytes reserved by the legacy atom.</summary>
        public ushort RawUnused { get; }

        /// <summary>Gets whether an inline SoundContainer overrides <see cref="SoundIdReference"/>.</summary>
        public bool HasSoundOverride { get; }

        private bool ReadFlag(int shift) => ((RawFlags >> shift) & 0x03U) == 1U;
    }
}
