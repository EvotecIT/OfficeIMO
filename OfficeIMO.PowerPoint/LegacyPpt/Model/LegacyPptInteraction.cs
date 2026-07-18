namespace OfficeIMO.PowerPoint.LegacyPpt.Model {
    /// <summary>Identifies how an interactive action is triggered.</summary>
    public enum LegacyPptInteractionTrigger : ushort {
        /// <summary>The action runs when the object or text is clicked.</summary>
        MouseClick = 0,

        /// <summary>The action runs when the pointer moves over the object or text.</summary>
        MouseOver = 1
    }

    /// <summary>Identifies the operation stored in an InteractiveInfoAtom.</summary>
    public enum LegacyPptInteractionAction : byte {
        /// <summary>No action.</summary>
        None = 0,
        /// <summary>Run a macro.</summary>
        Macro = 1,
        /// <summary>Run an external program.</summary>
        RunProgram = 2,
        /// <summary>Jump within the slide show.</summary>
        Jump = 3,
        /// <summary>Open a hyperlink.</summary>
        Hyperlink = 4,
        /// <summary>Invoke an OLE verb.</summary>
        Ole = 5,
        /// <summary>Play media.</summary>
        Media = 6,
        /// <summary>Open a named custom show.</summary>
        CustomShow = 7
    }

    /// <summary>Identifies a built-in slide-show jump.</summary>
    public enum LegacyPptInteractionJump : byte {
        /// <summary>No built-in jump.</summary>
        None = 0,
        /// <summary>Next slide.</summary>
        NextSlide = 1,
        /// <summary>Previous slide.</summary>
        PreviousSlide = 2,
        /// <summary>First slide.</summary>
        FirstSlide = 3,
        /// <summary>Last slide.</summary>
        LastSlide = 4,
        /// <summary>Last viewed slide.</summary>
        LastViewedSlide = 5,
        /// <summary>End the slide show.</summary>
        EndShow = 6
    }

    /// <summary>Identifies the legacy hyperlink target category.</summary>
    public enum LegacyPptHyperlinkType : byte {
        /// <summary>Next slide.</summary>
        NextSlide = 0,
        /// <summary>Previous slide.</summary>
        PreviousSlide = 1,
        /// <summary>First slide.</summary>
        FirstSlide = 2,
        /// <summary>Last slide.</summary>
        LastSlide = 3,
        /// <summary>Named custom show.</summary>
        CustomShow = 6,
        /// <summary>A slide number or named slide.</summary>
        SlideNumber = 7,
        /// <summary>An Internet URL.</summary>
        Url = 8,
        /// <summary>Another presentation.</summary>
        OtherPresentation = 9,
        /// <summary>Another file.</summary>
        OtherFile = 10,
        /// <summary>No hyperlink category.</summary>
        Nil = 255
    }

    /// <summary>Represents one document-level external or internal-slide hyperlink target.</summary>
    public sealed class LegacyPptHyperlink {
        internal LegacyPptHyperlink(uint id, string? friendlyName, string? target,
            string? location) {
            Id = id;
            FriendlyName = friendlyName;
            Target = target;
            Location = location;
            if (string.IsNullOrEmpty(target)
                && TryParseSlideLocation(location, out uint slideId,
                    out int slideNumber, out string? slideName)) {
                TargetSlideId = slideId;
                TargetSlideNumber = slideNumber;
                TargetSlideName = slideName;
            }
        }

        internal void ApplyExtension(string? screenTip, uint flags) {
            ScreenTip = screenTip;
            ExtensionFlags = flags;
        }

        /// <summary>Gets the identifier referenced by InteractiveInfoAtom records.</summary>
        public uint Id { get; }

        /// <summary>Gets the user-readable target name, when present.</summary>
        public string? FriendlyName { get; }

        /// <summary>Gets the destination path or URL, when present.</summary>
        public string? Target { get; }

        /// <summary>Gets the location within the destination, when present.</summary>
        public string? Location { get; }

        /// <summary>Gets the binary slide identifier for an internal-slide target, when present.</summary>
        public uint? TargetSlideId { get; }

        /// <summary>Gets the one-based slide ordinal recorded with an internal-slide target, when present.</summary>
        public int? TargetSlideNumber { get; }

        /// <summary>Gets the recorded internal-slide name, when present.</summary>
        public string? TargetSlideName { get; }

        /// <summary>Gets whether this hyperlink points to a slide in the same presentation.</summary>
        public bool IsInternalSlideTarget => TargetSlideId.HasValue;

        /// <summary>Gets the PowerPoint 2000+ screen tip, when present.</summary>
        public string? ScreenTip { get; private set; }

        /// <summary>Gets the raw PowerPoint 2000+ hyperlink flags.</summary>
        public uint ExtensionFlags { get; private set; }

        /// <summary>Gets whether the hyperlink was created through the Insert Hyperlink dialog.</summary>
        public bool WasCreatedByInsertHyperlinkDialog => (ExtensionFlags & 0x01U) != 0;

        /// <summary>Gets whether the location identifies a named custom show.</summary>
        public bool LocationIsNamedShow => (ExtensionFlags & 0x02U) != 0;

        /// <summary>Gets whether a named custom show returns to the invoking slide.</summary>
        public bool ReturnsToSlideAfterCustomShow => (ExtensionFlags & 0x04U) != 0;

        /// <summary>Gets a combined URI when the target can be represented as one.</summary>
        public Uri? Uri {
            get {
                if (IsInternalSlideTarget) return null;
                string value = Target ?? string.Empty;
                if (!string.IsNullOrEmpty(Location)) {
                    value = string.IsNullOrEmpty(value)
                        ? Location!
                        : value + (value.IndexOf("#", StringComparison.Ordinal) >= 0
                            ? string.Empty
                            : "#") + Location;
                }
                return Uri.TryCreate(value, UriKind.RelativeOrAbsolute, out Uri? uri) ? uri : null;
            }
        }

        private static bool TryParseSlideLocation(string? value, out uint slideId,
            out int slideNumber, out string? slideName) {
            slideId = 0;
            slideNumber = 0;
            slideName = null;
            if (value == null || value.Length == 0) return false;
            string location = value;
            int firstComma = location.IndexOf(',');
            int secondComma = firstComma < 0 ? -1 : location.IndexOf(',', firstComma + 1);
            if (firstComma <= 0 || secondComma <= firstComma + 1
                || !uint.TryParse(location.Substring(0, firstComma).Trim(), out slideId)
                || slideId < 256U || slideId > 0x7FFFFFFFU
                || !int.TryParse(location.Substring(firstComma + 1,
                    secondComma - firstComma - 1).Trim(), out slideNumber)
                || slideNumber <= 0) {
                slideId = 0;
                slideNumber = 0;
                return false;
            }
            string name = location.Substring(secondComma + 1).Trim();
            slideName = name.Length == 0 ? null : name;
            return true;
        }
    }

    /// <summary>Represents a decoded click or mouse-over action.</summary>
    public sealed class LegacyPptInteraction {
        internal LegacyPptInteraction(LegacyPptInteractionTrigger trigger,
            LegacyPptInteractionAction action, LegacyPptInteractionJump jump,
            LegacyPptHyperlinkType hyperlinkType, uint soundIdReference,
            uint hyperlinkIdReference, byte oleVerb, byte flags, string? name,
            LegacyPptHyperlink? hyperlink, LegacyPptCustomShow? customShow) {
            Trigger = trigger;
            Action = action;
            Jump = jump;
            HyperlinkType = hyperlinkType;
            SoundIdReference = soundIdReference;
            HyperlinkIdReference = hyperlinkIdReference;
            OleVerb = oleVerb;
            Flags = flags;
            Name = name;
            Hyperlink = hyperlink;
            CustomShow = customShow;
        }

        /// <summary>Gets the interaction trigger.</summary>
        public LegacyPptInteractionTrigger Trigger { get; }

        /// <summary>Gets the action kind.</summary>
        public LegacyPptInteractionAction Action { get; }

        /// <summary>Gets the built-in jump kind.</summary>
        public LegacyPptInteractionJump Jump { get; }

        /// <summary>Gets the legacy hyperlink target category.</summary>
        public LegacyPptHyperlinkType HyperlinkType { get; }

        /// <summary>Gets the referenced transition or action sound identifier.</summary>
        public uint SoundIdReference { get; }

        /// <summary>Gets the referenced document hyperlink identifier.</summary>
        public uint HyperlinkIdReference { get; }

        /// <summary>Gets the OLE verb value.</summary>
        public byte OleVerb { get; }

        /// <summary>Gets the raw animated, stop-sound, custom-show-return, and visited flags.</summary>
        public byte Flags { get; }

        /// <summary>Gets the macro, program, or custom-show name, when present.</summary>
        public string? Name { get; }

        /// <summary>Gets the resolved document-level hyperlink, when present.</summary>
        public LegacyPptHyperlink? Hyperlink { get; }

        /// <summary>Gets the resolved named show for a custom-show action, when present.</summary>
        public LegacyPptCustomShow? CustomShow { get; }

        /// <summary>Gets whether the action is marked as animated.</summary>
        public bool IsAnimated => (Flags & 0x01) != 0;

        /// <summary>Gets whether the action stops the current sound.</summary>
        public bool StopsSound => (Flags & 0x02) != 0;

        /// <summary>Gets whether a custom show returns to the invoking slide.</summary>
        public bool ReturnsFromCustomShow => (Flags & 0x04) != 0;

        /// <summary>Gets whether PowerPoint marked the interaction as visited.</summary>
        public bool IsVisited => (Flags & 0x08) != 0;
    }

    /// <summary>Anchors an interaction to a half-open range of shape text.</summary>
    public sealed class LegacyPptTextInteraction {
        internal LegacyPptTextInteraction(int start, int length,
            LegacyPptInteraction interaction) {
            Start = start;
            Length = length;
            Interaction = interaction ?? throw new ArgumentNullException(nameof(interaction));
        }

        /// <summary>Gets the zero-based start in the normalized text.</summary>
        public int Start { get; }

        /// <summary>Gets the number of characters covered by the interaction.</summary>
        public int Length { get; }

        /// <summary>Gets the decoded action.</summary>
        public LegacyPptInteraction Interaction { get; }
    }
}
