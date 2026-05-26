namespace OfficeIMO.Visio.Diagrams {
    /// <summary>
    /// Semantic component kinds used by the architecture diagram builder.
    /// </summary>
    public enum VisioArchitectureShapeKind {
        /// <summary>Human or external actor.</summary>
        Actor,

        /// <summary>Application, service, or process component.</summary>
        Service,

        /// <summary>Compute host, VM, container, or worker.</summary>
        Compute,

        /// <summary>Gateway, ingress, load balancer, or public endpoint.</summary>
        Gateway,

        /// <summary>Database or structured data store.</summary>
        Database,

        /// <summary>Blob/file/object storage.</summary>
        Storage,

        /// <summary>Queue, bus, stream, or asynchronous broker.</summary>
        Queue,

        /// <summary>Security boundary, identity, key, or policy component.</summary>
        Security,

        /// <summary>Network, subnet, or routing component.</summary>
        Network,

        /// <summary>External dependency or third-party system.</summary>
        External
    }
}
