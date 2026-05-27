namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Semantic node kinds used by the network diagram builder.
    /// </summary>
    public enum VisioNetworkNodeKind {
        /// <summary>External user or client.</summary>
        User,

        /// <summary>Desktop, laptop, or endpoint device.</summary>
        Workstation,

        /// <summary>Server or virtual machine.</summary>
        Server,

        /// <summary>Network switch.</summary>
        Switch,

        /// <summary>Router or routing appliance.</summary>
        Router,

        /// <summary>Firewall or security boundary.</summary>
        Firewall,

        /// <summary>Internet or external network.</summary>
        Internet,

        /// <summary>Printer or peripheral.</summary>
        Printer,

        /// <summary>Storage appliance.</summary>
        Storage,

        /// <summary>Database node.</summary>
        Database,

        /// <summary>Wireless access point.</summary>
        Wireless,

        /// <summary>Legend, note, or annotation.</summary>
        Note
    }
}
