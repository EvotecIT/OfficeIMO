using System;

namespace OfficeIMO;

/// <summary>Controls how an existing output path is handled.</summary>
public enum OfficePublicationPolicy {
    /// <summary>Fail before publishing when the destination exists.</summary>
    Fail,
    /// <summary>Atomically replace the destination after the staged artifact is validated.</summary>
    Replace,
    /// <summary>Publish to the first available numbered file name.</summary>
    Rename
}
