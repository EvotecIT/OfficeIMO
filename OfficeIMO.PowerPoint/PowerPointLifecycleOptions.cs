namespace OfficeIMO.PowerPoint {
    /// <summary>Controls whether an existing presentation is opened for editing or inspection.</summary>
    public enum PowerPointOpenMode {
        /// <summary>The presentation is editable and file-backed changes are saved on dispose.</summary>
        Edit,
        /// <summary>The presentation is opened without write access or repair behavior.</summary>
        ReadOnly
    }

    /// <summary>Controls persistence for a presentation created on a caller-owned stream.</summary>
    public sealed class PowerPointStreamCreateOptions {
        /// <summary>Writes the completed package back to the stream when the presentation is disposed.</summary>
        public bool AutoSave { get; set; } = true;
    }

    /// <summary>Controls access and persistence for a presentation opened from a caller-owned stream.</summary>
    public sealed class PowerPointStreamOpenOptions {
        /// <summary>Opens the presentation for editing or read-only inspection.</summary>
        public PowerPointOpenMode Mode { get; set; } = PowerPointOpenMode.Edit;

        /// <summary>Writes editable changes back to the source stream when the presentation is disposed.</summary>
        public bool AutoSave { get; set; }
    }
}
