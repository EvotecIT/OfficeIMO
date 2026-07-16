namespace OfficeIMO.PowerPoint {
    /// <summary>Specifies a classic PowerPoint entrance effect representable by binary PPT.</summary>
    public enum PowerPointClassicAnimationEffect : byte {
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

    /// <summary>Specifies how a shape, text body, or chart is built.</summary>
    public enum PowerPointClassicAnimationBuildType : byte {
        /// <summary>No animation build.</summary>
        None = 0x00,
        /// <summary>Animate as one object.</summary>
        AsOneObject = 0x01,
        /// <summary>Animate paragraphs through level 1 separately.</summary>
        ByLevel1Paragraph = 0x02,
        /// <summary>Animate paragraphs through level 2 separately.</summary>
        ByLevel2Paragraph = 0x03,
        /// <summary>Animate paragraphs through level 3 separately.</summary>
        ByLevel3Paragraph = 0x04,
        /// <summary>Animate paragraphs through level 4 separately.</summary>
        ByLevel4Paragraph = 0x05,
        /// <summary>Animate paragraphs through level 5 separately.</summary>
        ByLevel5Paragraph = 0x06,
        /// <summary>Animate chart series separately.</summary>
        ChartBySeries = 0x07,
        /// <summary>Animate chart categories separately.</summary>
        ChartByCategory = 0x08,
        /// <summary>Animate chart elements in series order.</summary>
        ChartByElementInSeries = 0x09,
        /// <summary>Animate chart elements in category order.</summary>
        ChartByElementInCategory = 0x0A,
        /// <summary>Follow the corresponding master placeholder.</summary>
        FollowMaster = 0xFE
    }

    /// <summary>Specifies the state applied after a classic animation.</summary>
    public enum PowerPointClassicAnimationAfterEffect : byte {
        /// <summary>No after-effect.</summary>
        None = 0x00,
        /// <summary>Dim to the authored color.</summary>
        Dim = 0x01,
        /// <summary>Hide on the next click.</summary>
        HideOnNextClick = 0x02,
        /// <summary>Hide immediately.</summary>
        HideImmediately = 0x03
    }

    /// <summary>Specifies how text is subdivided within a classic animation.</summary>
    public enum PowerPointClassicTextBuild : byte {
        /// <summary>Animate the text at once.</summary>
        AllAtOnce = 0x00,
        /// <summary>Animate word by word.</summary>
        ByWord = 0x01,
        /// <summary>Animate character by character.</summary>
        ByCharacter = 0x02
    }

    /// <summary>Options used when authoring a classic shape or text animation.</summary>
    public sealed class PowerPointClassicAnimationOptions {
        /// <summary>Gets or sets the effect-specific direction or variant byte.</summary>
        public byte Direction { get; set; }

        /// <summary>Gets or sets the shape, text, or chart build mode.</summary>
        public PowerPointClassicAnimationBuildType BuildType { get; set; } =
            PowerPointClassicAnimationBuildType.AsOneObject;

        /// <summary>Gets or sets whether the effect starts automatically.</summary>
        public bool Automatic { get; set; }

        /// <summary>Gets or sets the automatic-start delay in milliseconds.</summary>
        public int DelayMilliseconds { get; set; }

        /// <summary>Gets or sets whether the build order is reversed.</summary>
        public bool Reverse { get; set; }

        /// <summary>Gets or sets whether the shape background participates in the effect.</summary>
        public bool AnimateBackground { get; set; }

        /// <summary>Gets or sets whether existing sounds stop when the effect starts.</summary>
        public bool StopsSound { get; set; }

        /// <summary>Gets or sets the state applied after the effect completes.</summary>
        public PowerPointClassicAnimationAfterEffect AfterEffect { get; set; }

        /// <summary>Gets or sets the word or character subdivision for text builds.</summary>
        public PowerPointClassicTextBuild TextBuild { get; set; }

        /// <summary>Gets or sets the raw legacy dim ColorIndexStruct.</summary>
        public uint RawDimColor { get; set; }
    }

    /// <summary>Describes one editable classic animation attached to a slide shape.</summary>
    public sealed class PowerPointClassicAnimation {
        internal PowerPointClassicAnimation(uint shapeId, PowerPointClassicAnimationEffect effect,
            byte direction, PowerPointClassicAnimationBuildType buildType, bool automatic,
            int delayMilliseconds, int order, bool reverse, bool animateBackground,
            PowerPointClassicAnimationAfterEffect afterEffect,
            PowerPointClassicTextBuild textBuild, uint rawDimColor,
            bool playsSound = false, bool stopsSound = false,
            string? soundRelationshipId = null, string? soundName = null) {
            ShapeId = shapeId;
            Effect = effect;
            Direction = direction;
            BuildType = buildType;
            Automatic = automatic;
            DelayMilliseconds = delayMilliseconds;
            Order = order;
            Reverse = reverse;
            AnimateBackground = animateBackground;
            AfterEffect = afterEffect;
            TextBuild = textBuild;
            RawDimColor = rawDimColor;
            PlaysSound = playsSound;
            StopsSound = stopsSound;
            SoundRelationshipId = soundRelationshipId;
            SoundName = soundName;
        }

        /// <summary>Gets the target shape identifier.</summary>
        public uint ShapeId { get; }

        /// <summary>Gets the classic entrance effect.</summary>
        public PowerPointClassicAnimationEffect Effect { get; }

        /// <summary>Gets the effect-specific direction or variant byte.</summary>
        public byte Direction { get; }

        /// <summary>Gets the shape, text, or chart build mode.</summary>
        public PowerPointClassicAnimationBuildType BuildType { get; }

        /// <summary>Gets whether the effect starts automatically.</summary>
        public bool Automatic { get; }

        /// <summary>Gets the automatic-start delay in milliseconds.</summary>
        public int DelayMilliseconds { get; }

        /// <summary>Gets the authored slide animation order; -2 follows the master placeholder.</summary>
        public int Order { get; }

        /// <summary>Gets whether the build order is reversed.</summary>
        public bool Reverse { get; }

        /// <summary>Gets whether the shape background participates in the effect.</summary>
        public bool AnimateBackground { get; }

        /// <summary>Gets the state applied after the effect completes.</summary>
        public PowerPointClassicAnimationAfterEffect AfterEffect { get; }

        /// <summary>Gets the text subdivision mode.</summary>
        public PowerPointClassicTextBuild TextBuild { get; }

        /// <summary>Gets the raw legacy dim ColorIndexStruct.</summary>
        public uint RawDimColor { get; }

        /// <summary>Gets whether an embedded sound plays with the effect.</summary>
        public bool PlaysSound { get; }

        /// <summary>Gets whether existing sounds stop when the effect starts.</summary>
        public bool StopsSound { get; }

        /// <summary>Gets the embedded-audio relationship used by the effect.</summary>
        public string? SoundRelationshipId { get; }

        /// <summary>Gets the embedded sound name.</summary>
        public string? SoundName { get; }
    }
}
