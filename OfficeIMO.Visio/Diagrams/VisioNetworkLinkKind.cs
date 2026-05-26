namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Semantic link kinds used by the network diagram builder.
    /// </summary>
    public enum VisioNetworkLinkKind {
        /// <summary>Standard network/data connection.</summary>
        Ethernet,

        /// <summary>Higher-capacity uplink or trunk.</summary>
        Trunk,

        /// <summary>Wireless connection.</summary>
        Wireless,

        /// <summary>Management or control-plane connection.</summary>
        Management
    }
}
