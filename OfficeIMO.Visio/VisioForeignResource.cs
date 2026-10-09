using System;
using System.Collections.Generic;

namespace OfficeIMO.Visio;

// Immutable binary content follows shape trees; page/master stores also retain opaque XML resources.
internal sealed class VisioForeignResource {
    internal string RelationshipId { get; set; } = string.Empty;
    internal string RelationshipType { get; set; } = string.Empty;
    internal string ContentType { get; set; } = string.Empty;
    internal byte[] Bytes { get; set; } = Array.Empty<byte>();
}

public partial class VisioPage {
    internal List<VisioForeignResource> ForeignResources { get; } = new();
}
