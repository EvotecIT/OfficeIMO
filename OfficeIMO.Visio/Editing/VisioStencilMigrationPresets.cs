using System;
using System.Globalization;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio {
    /// <summary>
    /// Reusable stencil migration presets for common loaded-diagram cleanup workflows.
    /// </summary>
    public static class VisioStencilMigrationPresets {
        /// <summary>
        /// Creates a migration map that upgrades unstenciled/basic flowchart-like masters to semantically matching catalog stencils.
        /// </summary>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.Flowchart"/> or a package-backed flowchart catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMap BasicFlowchart(VisioStencilCatalog catalog, bool resizeToStencil = true) {
            return VisioStencilMigrationMap.Create(builder => builder.MapBasicFlowchart(catalog, resizeToStencil));
        }

        /// <summary>
        /// Creates a migration map that upgrades common unstenciled network and infrastructure shapes to matching catalog stencils.
        /// </summary>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.Network"/>, <see cref="VisioStencils.Infrastructure"/>, or a package-backed infrastructure catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMap NetworkInfrastructure(VisioStencilCatalog catalog, bool resizeToStencil = true) {
            return VisioStencilMigrationMap.Create(builder => builder.MapNetworkInfrastructure(catalog, resizeToStencil));
        }

        /// <summary>
        /// Creates a migration map that upgrades common unstenciled architecture shapes to matching catalog stencils.
        /// </summary>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.Architecture"/>, <see cref="VisioStencils.Cloud"/>, or a package-backed architecture catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMap ArchitectureInfrastructure(VisioStencilCatalog catalog, bool resizeToStencil = true) {
            return VisioStencilMigrationMap.Create(builder => builder.MapArchitectureInfrastructure(catalog, resizeToStencil));
        }

        /// <summary>
        /// Creates a migration map that upgrades common unstenciled organization chart shapes to matching catalog stencils.
        /// </summary>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.OrgChart"/> or a package-backed organization chart catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMap OrgChart(VisioStencilCatalog catalog, bool resizeToStencil = true) {
            return VisioStencilMigrationMap.Create(builder => builder.MapOrgChart(catalog, resizeToStencil));
        }

        /// <summary>
        /// Creates a migration map that upgrades common unstenciled timeline and roadmap shapes to matching catalog stencils.
        /// </summary>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.Timeline"/> or a package-backed timeline catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMap Timeline(VisioStencilCatalog catalog, bool resizeToStencil = true) {
            return VisioStencilMigrationMap.Create(builder => builder.MapTimeline(catalog, resizeToStencil));
        }

        /// <summary>
        /// Creates a migration map that upgrades common unstenciled sequence diagram shapes to matching catalog stencils.
        /// </summary>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.Sequence"/> or a package-backed sequence diagram catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMap Sequence(VisioStencilCatalog catalog, bool resizeToStencil = true) {
            return VisioStencilMigrationMap.Create(builder => builder.MapSequence(catalog, resizeToStencil));
        }

        /// <summary>
        /// Creates a migration map that upgrades common unstenciled swimlane and process-map shapes to matching catalog stencils.
        /// </summary>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.Swimlane"/> or a package-backed process-map catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMap SwimlaneProcessMap(VisioStencilCatalog catalog, bool resizeToStencil = true) {
            return VisioStencilMigrationMap.Create(builder => builder.MapSwimlaneProcessMap(catalog, resizeToStencil));
        }

        /// <summary>
        /// Creates a migration map that upgrades common unstenciled cloud infrastructure shapes to matching catalog stencils.
        /// </summary>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.Cloud"/> or a package-backed cloud catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMap CloudInfrastructure(VisioStencilCatalog catalog, bool resizeToStencil = true) {
            return VisioStencilMigrationMap.Create(builder => builder.MapCloudInfrastructure(catalog, resizeToStencil));
        }

        /// <summary>
        /// Creates a migration map that upgrades common unstenciled security and identity shapes to matching catalog stencils.
        /// </summary>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.SecurityIdentity"/> or a package-backed security catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMap SecurityIdentity(VisioStencilCatalog catalog, bool resizeToStencil = true) {
            return VisioStencilMigrationMap.Create(builder => builder.MapSecurityIdentity(catalog, resizeToStencil));
        }

        /// <summary>
        /// Creates a migration map that upgrades common unstenciled Kubernetes and container-platform shapes to matching catalog stencils.
        /// </summary>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.ContainersKubernetes"/> or a package-backed container platform catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMap ContainersKubernetes(VisioStencilCatalog catalog, bool resizeToStencil = true) {
            return VisioStencilMigrationMap.Create(builder => builder.MapContainersKubernetes(catalog, resizeToStencil));
        }

        /// <summary>
        /// Creates a migration map that upgrades common unstenciled data and platform-service shapes to matching catalog stencils.
        /// </summary>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.DataPlatform"/> or a package-backed data platform catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMap DataPlatform(VisioStencilCatalog catalog, bool resizeToStencil = true) {
            return VisioStencilMigrationMap.Create(builder => builder.MapDataPlatform(catalog, resizeToStencil));
        }

        /// <summary>
        /// Creates a migration map that upgrades common unstenciled collaboration and business-process shapes to matching catalog stencils.
        /// </summary>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.CollaborationBusiness"/> or a package-backed business-process catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMap CollaborationBusiness(VisioStencilCatalog catalog, bool resizeToStencil = true) {
            return VisioStencilMigrationMap.Create(builder => builder.MapCollaborationBusiness(catalog, resizeToStencil));
        }

        /// <summary>
        /// Adds rules that upgrade unstenciled/basic flowchart-like masters to semantically matching catalog stencils.
        /// Already stenciled shapes carrying OfficeIMO stencil metadata are not matched.
        /// </summary>
        /// <param name="builder">Migration map builder.</param>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.Flowchart"/> or a package-backed flowchart catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMapBuilder MapBasicFlowchart(this VisioStencilMigrationMapBuilder builder, VisioStencilCatalog catalog, bool resizeToStencil = true) {
            if (builder == null) {
                throw new ArgumentNullException(nameof(builder));
            }

            if (catalog == null) {
                throw new ArgumentNullException(nameof(catalog));
            }

            int mapped = 0;
            MapIfFound(builder, "Process", catalog, new[] { "process", "task", "step" }, resizeToStencil, ref mapped);
            MapIfFound(builder, "Rectangle", catalog, new[] { "process", "task", "step" }, resizeToStencil, ref mapped);
            MapIfFound(builder, "Decision", catalog, new[] { "decision", "branch", "choice" }, resizeToStencil, ref mapped);
            MapIfFound(builder, "Diamond", catalog, new[] { "decision", "branch", "choice" }, resizeToStencil, ref mapped);
            MapIfFound(builder, "Data", catalog, new[] { "data", "input", "output" }, resizeToStencil, ref mapped);
            MapIfFound(builder, "Parallelogram", catalog, new[] { "data", "input", "output" }, resizeToStencil, ref mapped);
            MapIfFound(builder, "Preparation", catalog, new[] { "preparation", "setup" }, resizeToStencil, ref mapped);
            MapIfFound(builder, "Hexagon", catalog, new[] { "preparation", "setup" }, resizeToStencil, ref mapped);
            MapIfFound(builder, "Ellipse", catalog, new[] { "start", "start/end", "terminator" }, resizeToStencil, ref mapped);
            MapIfFound(builder, "Manual operation", catalog, new[] { "manual operation", "manual" }, resizeToStencil, ref mapped);
            MapIfFound(builder, "Off-page reference", catalog, new[] { "off-page reference", "off-page" }, resizeToStencil, ref mapped);

            if (mapped == 0) {
                throw new InvalidOperationException($"Catalog '{catalog.Name}' does not contain any flowchart-compatible stencil shapes.");
            }

            return builder;
        }

        /// <summary>
        /// Adds rules that upgrade common unstenciled network and infrastructure shapes to semantically matching catalog stencils.
        /// Already stenciled shapes carrying OfficeIMO stencil metadata are not matched.
        /// </summary>
        /// <param name="builder">Migration map builder.</param>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.Network"/>, <see cref="VisioStencils.Infrastructure"/>, or a package-backed infrastructure catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMapBuilder MapNetworkInfrastructure(this VisioStencilMigrationMapBuilder builder, VisioStencilCatalog catalog, bool resizeToStencil = true) {
            if (builder == null) {
                throw new ArgumentNullException(nameof(builder));
            }

            if (catalog == null) {
                throw new ArgumentNullException(nameof(catalog));
            }

            int mapped = 0;
            MapIfFound(builder, UnstenciledIdentity("firewall", "waf", "acl"), catalog, new[] { "firewall", "network-security", "edge" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("load balancer", "load-balancer", "lb", "proxy"), catalog, new[] { "load-balancer", "traffic", "proxy", "gateway", "router" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("router", "gateway", "route"), catalog, new[] { "router", "gateway", "route" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("switch", "hub", "lan"), catalog, new[] { "switch", "hub", "lan" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("wireless", "wifi", "access point", "access-point"), catalog, new[] { "wireless", "access-point", "wifi" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("internet", "wan", "cloud"), catalog, new[] { "internet", "wan", "cloud" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("database", "sql", "db"), catalog, new[] { "database", "sql" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("storage", "nas", "san", "backup"), catalog, new[] { "storage", "nas", "san", "backup" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("rack", "cabinet", "datacenter"), catalog, new[] { "rack", "cabinet", "datacenter" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("printer", "peripheral"), catalog, new[] { "printer", "peripheral" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("workstation", "desktop", "laptop", "endpoint", "client"), catalog, new[] { "workstation", "desktop", "laptop", "endpoint" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("server", "host", "vm", "compute", "node", "appliance"), catalog, new[] { "server", "host", "compute", "node", "appliance" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("zone", "subnet", "boundary", "vlan"), catalog, new[] { "zone", "subnet", "boundary", "network" }, resizeToStencil, ref mapped);

            if (mapped == 0) {
                throw new InvalidOperationException($"Catalog '{catalog.Name}' does not contain any network or infrastructure-compatible stencil shapes.");
            }

            return builder;
        }

        /// <summary>
        /// Adds rules that upgrade common unstenciled architecture shapes to semantically matching catalog stencils.
        /// Already stenciled shapes carrying OfficeIMO stencil metadata are not matched.
        /// </summary>
        /// <param name="builder">Migration map builder.</param>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.Architecture"/>, <see cref="VisioStencils.Cloud"/>, or a package-backed architecture catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMapBuilder MapArchitectureInfrastructure(this VisioStencilMigrationMapBuilder builder, VisioStencilCatalog catalog, bool resizeToStencil = true) {
            if (builder == null) {
                throw new ArgumentNullException(nameof(builder));
            }

            if (catalog == null) {
                throw new ArgumentNullException(nameof(catalog));
            }

            int mapped = 0;
            MapIfFound(builder, UnstenciledIdentity("gateway", "ingress", "endpoint", "api gateway", "api-gateway"), catalog, new[] { "gateway", "ingress", "endpoint" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("database", "sql", "data store", "data-store", "db"), catalog, new[] { "database", "sql", "data-store" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("storage", "blob", "file", "object store", "object-store"), catalog, new[] { "storage", "blob", "object-store" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("queue", "bus", "broker", "stream", "event"), catalog, new[] { "queue", "bus", "broker" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("security", "identity", "policy", "key", "secret"), catalog, new[] { "security", "identity", "policy", "key" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("actor", "user", "person", "client"), catalog, new[] { "actor", "user", "person", "client" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("external", "third party", "third-party", "partner", "vendor"), catalog, new[] { "external", "third-party", "partner" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("network", "subnet", "vnet", "route"), catalog, new[] { "network", "subnet", "vnet", "route" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("region", "zone", "boundary", "container"), catalog, new[] { "region", "zone", "boundary", "container" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("service", "application", "app", "api", "platform"), catalog, new[] { "service", "application", "app", "api" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("compute", "server", "worker", "vm", "container"), catalog, new[] { "compute", "server", "worker", "container" }, resizeToStencil, ref mapped);

            if (mapped == 0) {
                throw new InvalidOperationException($"Catalog '{catalog.Name}' does not contain any architecture-compatible stencil shapes.");
            }

            return builder;
        }

        /// <summary>
        /// Adds rules that upgrade common unstenciled organization chart shapes to semantically matching catalog stencils.
        /// Already stenciled shapes carrying OfficeIMO stencil metadata are not matched.
        /// </summary>
        /// <param name="builder">Migration map builder.</param>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.OrgChart"/> or a package-backed organization chart catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMapBuilder MapOrgChart(this VisioStencilMigrationMapBuilder builder, VisioStencilCatalog catalog, bool resizeToStencil = true) {
            if (builder == null) {
                throw new ArgumentNullException(nameof(builder));
            }

            if (catalog == null) {
                throw new ArgumentNullException(nameof(catalog));
            }

            int mapped = 0;
            MapIfFound(builder, UnstenciledIdentity("executive assistant", "assistant"), catalog, new[] { "assistant", "ea", "staff" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("executive", "ceo", "cto", "cfo", "cio", "chief", "leader"), catalog, new[] { "executive", "root", "leader", "ceo" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("manager", "lead", "supervisor"), catalog, new[] { "manager", "lead", "supervisor" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("vacancy", "open role", "open position", "hiring", "backfill"), catalog, new[] { "vacancy", "open", "hiring" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("external", "advisor", "vendor", "partner", "consultant"), catalog, new[] { "external", "advisor", "vendor", "partner" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("team band", "department", "team", "group", "org unit"), catalog, new[] { "team", "department", "container", "group" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("employee", "person", "role", "engineer", "analyst", "specialist", "contributor"), catalog, new[] { "position", "person", "employee", "role" }, resizeToStencil, ref mapped);

            if (mapped == 0) {
                throw new InvalidOperationException($"Catalog '{catalog.Name}' does not contain any organization chart-compatible stencil shapes.");
            }

            return builder;
        }

        /// <summary>
        /// Adds rules that upgrade common unstenciled timeline and roadmap shapes to semantically matching catalog stencils.
        /// Already stenciled shapes carrying OfficeIMO stencil metadata are not matched.
        /// </summary>
        /// <param name="builder">Migration map builder.</param>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.Timeline"/> or a package-backed timeline catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMapBuilder MapTimeline(this VisioStencilMigrationMapBuilder builder, VisioStencilCatalog catalog, bool resizeToStencil = true) {
            if (builder == null) {
                throw new ArgumentNullException(nameof(builder));
            }

            if (catalog == null) {
                throw new ArgumentNullException(nameof(catalog));
            }

            int mapped = 0;
            MapIfFound(builder, UnstenciledIdentity("timeline axis", "roadmap axis", "schedule axis"), catalog, new[] { "axis", "timeline", "roadmap", "schedule" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("release", "launch", "delivery"), catalog, new[] { "release", "delivery", "launch" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("decision", "approval", "gate"), catalog, new[] { "decision", "approval", "gate" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("risk", "issue", "attention"), catalog, new[] { "risk", "issue", "attention" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("milestone", "date marker", "marker"), catalog, new[] { "milestone", "date", "marker" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("phase", "workstream", "duration", "span"), catalog, new[] { "span", "phase", "workstream", "duration" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("label", "annotation", "callout"), catalog, new[] { "label", "annotation", "callout" }, resizeToStencil, ref mapped);

            if (mapped == 0) {
                throw new InvalidOperationException($"Catalog '{catalog.Name}' does not contain any timeline-compatible stencil shapes.");
            }

            return builder;
        }

        /// <summary>
        /// Adds rules that upgrade common unstenciled sequence diagram shapes to semantically matching catalog stencils.
        /// Already stenciled shapes carrying OfficeIMO stencil metadata are not matched.
        /// </summary>
        /// <param name="builder">Migration map builder.</param>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.Sequence"/> or a package-backed sequence diagram catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMapBuilder MapSequence(this VisioStencilMigrationMapBuilder builder, VisioStencilCatalog catalog, bool resizeToStencil = true) {
            if (builder == null) {
                throw new ArgumentNullException(nameof(builder));
            }

            if (catalog == null) {
                throw new ArgumentNullException(nameof(catalog));
            }

            int mapped = 0;
            MapIfFound(builder, UnstenciledIdentity("actor", "user", "person", "client"), catalog, new[] { "actor", "user", "person", "client" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("boundary", "interface", "edge"), catalog, new[] { "boundary", "interface", "edge" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("control", "controller", "coordinator"), catalog, new[] { "control", "controller", "coordinator" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("entity", "domain object", "domain", "object"), catalog, new[] { "entity", "domain", "object" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("database", "data store", "data-store", "store"), catalog, new[] { "database", "store", "data-store" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("activation", "execution", "focus"), catalog, new[] { "activation", "execution", "focus" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("fragment", "combined fragment", "alt", "opt", "loop", "critical region"), catalog, new[] { "fragment", "combined-fragment", "alt", "opt", "loop" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("note", "annotation", "callout"), catalog, new[] { "note", "annotation", "callout" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("participant", "lifeline", "service", "component", "system"), catalog, new[] { "participant", "lifeline", "service", "component" }, resizeToStencil, ref mapped);

            if (mapped == 0) {
                throw new InvalidOperationException($"Catalog '{catalog.Name}' does not contain any sequence diagram-compatible stencil shapes.");
            }

            return builder;
        }

        /// <summary>
        /// Adds rules that upgrade common unstenciled swimlane and process-map shapes to semantically matching catalog stencils.
        /// Already stenciled shapes carrying OfficeIMO stencil metadata are not matched.
        /// </summary>
        /// <param name="builder">Migration map builder.</param>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.Swimlane"/> or a package-backed process-map catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMapBuilder MapSwimlaneProcessMap(this VisioStencilMigrationMapBuilder builder, VisioStencilCatalog catalog, bool resizeToStencil = true) {
            if (builder == null) {
                throw new ArgumentNullException(nameof(builder));
            }

            if (catalog == null) {
                throw new ArgumentNullException(nameof(catalog));
            }

            int mapped = 0;
            MapIfFound(builder, UnstenciledIdentity("lane", "swimlane", "responsibility lane", "participant lane"), catalog, new[] { "lane", "swimlane", "role", "participant" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("phase", "stage", "column"), catalog, new[] { "phase", "stage", "column", "milestone" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("decision", "branch", "choice"), catalog, new[] { "decision", "branch", "choice" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("input", "output", "document", "data"), catalog, new[] { "data", "input", "output", "document" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("start", "end", "terminator"), catalog, new[] { "start", "end", "terminator" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("activity", "task", "step", "process"), catalog, new[] { "activity", "task", "step", "process" }, resizeToStencil, ref mapped);

            if (mapped == 0) {
                throw new InvalidOperationException($"Catalog '{catalog.Name}' does not contain any swimlane or process-map-compatible stencil shapes.");
            }

            return builder;
        }

        /// <summary>
        /// Adds rules that upgrade common unstenciled cloud infrastructure shapes to semantically matching catalog stencils.
        /// Already stenciled shapes carrying OfficeIMO stencil metadata are not matched.
        /// </summary>
        /// <param name="builder">Migration map builder.</param>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.Cloud"/> or a package-backed cloud catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMapBuilder MapCloudInfrastructure(this VisioStencilMigrationMapBuilder builder, VisioStencilCatalog catalog, bool resizeToStencil = true) {
            if (builder == null) {
                throw new ArgumentNullException(nameof(builder));
            }

            if (catalog == null) {
                throw new ArgumentNullException(nameof(catalog));
            }

            int mapped = 0;
            MapIfFound(builder, UnstenciledIdentity("subscription", "tenant", "account boundary"), catalog, new[] { "subscription", "tenant", "account", "boundary" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("region", "availability zone", "location"), catalog, new[] { "region", "location", "zone" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("api gateway", "cloud gateway", "ingress"), catalog, new[] { "gateway", "api-gateway", "ingress" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("function", "serverless", "lambda"), catalog, new[] { "function", "serverless", "lambda" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("queue", "message queue", "event queue"), catalog, new[] { "queue", "event", "message" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("secret store", "key vault", "vault"), catalog, new[] { "secret-store", "secret", "vault", "key" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("monitoring", "observability", "metrics", "logs"), catalog, new[] { "monitoring", "observability", "metrics", "logs" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("cloud service", "managed service", "paas", "service"), catalog, new[] { "service", "managed-service", "paas" }, resizeToStencil, ref mapped);

            if (mapped == 0) {
                throw new InvalidOperationException($"Catalog '{catalog.Name}' does not contain any cloud infrastructure-compatible stencil shapes.");
            }

            return builder;
        }

        /// <summary>
        /// Adds rules that upgrade common unstenciled security and identity shapes to semantically matching catalog stencils.
        /// Already stenciled shapes carrying OfficeIMO stencil metadata are not matched.
        /// </summary>
        /// <param name="builder">Migration map builder.</param>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.SecurityIdentity"/> or a package-backed security catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMapBuilder MapSecurityIdentity(this VisioStencilMigrationMapBuilder builder, VisioStencilCatalog catalog, bool resizeToStencil = true) {
            if (builder == null) {
                throw new ArgumentNullException(nameof(builder));
            }

            if (catalog == null) {
                throw new ArgumentNullException(nameof(catalog));
            }

            int mapped = 0;
            MapIfFound(builder, UnstenciledIdentity("identity provider", "idp", "entra", "azure ad", "saml", "oidc"), catalog, new[] { "identity-provider", "idp", "entra", "saml", "oidc" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("user principal", "service account", "identity account"), catalog, new[] { "user", "account", "person", "identity" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("security group", "rbac group", "role group"), catalog, new[] { "group", "role", "rbac", "principal" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("firewall policy", "waf", "acl", "network security"), catalog, new[] { "firewall", "network-security", "waf", "acl" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("conditional access", "policy", "rule", "control"), catalog, new[] { "policy", "conditional-access", "rule", "control" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("certificate", "secret", "credential", "key vault"), catalog, new[] { "key", "certificate", "secret", "credential" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("audit log", "evidence", "compliance log"), catalog, new[] { "audit", "evidence", "compliance", "log" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("trust boundary", "security boundary", "threat model zone"), catalog, new[] { "trust-boundary", "boundary", "zone", "threat-model" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("security alert", "incident", "finding", "risk"), catalog, new[] { "alert", "incident", "finding", "risk" }, resizeToStencil, ref mapped);

            if (mapped == 0) {
                throw new InvalidOperationException($"Catalog '{catalog.Name}' does not contain any security or identity-compatible stencil shapes.");
            }

            return builder;
        }

        /// <summary>
        /// Adds rules that upgrade common unstenciled Kubernetes and container-platform shapes to semantically matching catalog stencils.
        /// Already stenciled shapes carrying OfficeIMO stencil metadata are not matched.
        /// </summary>
        /// <param name="builder">Migration map builder.</param>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.ContainersKubernetes"/> or a package-backed container platform catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMapBuilder MapContainersKubernetes(this VisioStencilMigrationMapBuilder builder, VisioStencilCatalog catalog, bool resizeToStencil = true) {
            if (builder == null) {
                throw new ArgumentNullException(nameof(builder));
            }

            if (catalog == null) {
                throw new ArgumentNullException(nameof(catalog));
            }

            int mapped = 0;
            MapIfFound(builder, UnstenciledIdentity("kubernetes cluster", "aks cluster", "eks cluster", "gke cluster", "cluster"), catalog, new[] { "cluster", "kubernetes", "aks", "eks", "gke" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("namespace", "tenant namespace", "scope"), catalog, new[] { "namespace", "tenant", "scope" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("worker node", "cluster node", "node host"), catalog, new[] { "node", "worker", "host" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("pod", "workload pod", "workload"), catalog, new[] { "pod", "workload" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("container image", "container runtime", "container"), catalog, new[] { "image", "runtime" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("kubernetes service", "k8s service", "svc", "service endpoint"), catalog, new[] { "service", "svc", "endpoint" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("ingress", "gateway route", "ingress route"), catalog, new[] { "ingress", "gateway", "route" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("config map", "configmap", "configuration", "settings"), catalog, new[] { "config", "configuration", "settings" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("secret", "credential", "key"), catalog, new[] { "secret", "credential", "key" }, resizeToStencil, ref mapped);

            if (mapped == 0) {
                throw new InvalidOperationException($"Catalog '{catalog.Name}' does not contain any Kubernetes or container-platform-compatible stencil shapes.");
            }

            return builder;
        }

        /// <summary>
        /// Adds rules that upgrade common unstenciled data and platform-service shapes to semantically matching catalog stencils.
        /// Already stenciled shapes carrying OfficeIMO stencil metadata are not matched.
        /// </summary>
        /// <param name="builder">Migration map builder.</param>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.DataPlatform"/> or a package-backed data platform catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMapBuilder MapDataPlatform(this VisioStencilMigrationMapBuilder builder, VisioStencilCatalog catalog, bool resizeToStencil = true) {
            if (builder == null) {
                throw new ArgumentNullException(nameof(builder));
            }

            if (catalog == null) {
                throw new ArgumentNullException(nameof(catalog));
            }

            int mapped = 0;
            MapIfFound(builder, UnstenciledIdentity("data lake", "analytics lake", "lake storage"), catalog, new[] { "lake", "analytics", "storage" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("warehouse", "data warehouse", "dwh", "mart"), catalog, new[] { "warehouse", "dwh", "mart" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("event stream", "stream", "kafka"), catalog, new[] { "stream", "event-stream", "kafka" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("pipeline", "etl", "elt", "job"), catalog, new[] { "pipeline", "etl", "elt", "job" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("data catalog", "metadata catalog", "lineage"), catalog, new[] { "catalog", "metadata", "lineage" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("data api", "query endpoint", "query api"), catalog, new[] { "api", "query", "endpoint" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("quality gate", "data quality", "validation", "dq"), catalog, new[] { "quality", "validation", "dq" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("sql database", "relational database", "database", "db"), catalog, new[] { "database", "sql", "relational" }, resizeToStencil, ref mapped);

            if (mapped == 0) {
                throw new InvalidOperationException($"Catalog '{catalog.Name}' does not contain any data or platform-compatible stencil shapes.");
            }

            return builder;
        }

        /// <summary>
        /// Adds rules that upgrade common unstenciled collaboration and business-process shapes to semantically matching catalog stencils.
        /// Already stenciled shapes carrying OfficeIMO stencil metadata are not matched.
        /// </summary>
        /// <param name="builder">Migration map builder.</param>
        /// <param name="catalog">Target catalog, usually <see cref="VisioStencils.CollaborationBusiness"/> or a package-backed business-process catalog.</param>
        /// <param name="resizeToStencil">Whether upgraded shapes should use the replacement stencil default size.</param>
        public static VisioStencilMigrationMapBuilder MapCollaborationBusiness(this VisioStencilMigrationMapBuilder builder, VisioStencilCatalog catalog, bool resizeToStencil = true) {
            if (builder == null) {
                throw new ArgumentNullException(nameof(builder));
            }

            if (catalog == null) {
                throw new ArgumentNullException(nameof(catalog));
            }

            int mapped = 0;
            MapIfFound(builder, UnstenciledIdentity("responsibility lane", "business lane", "swimlane", "owner lane"), catalog, new[] { "lane", "owner", "role", "swimlane" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("approval", "sign-off", "sign off", "business review"), catalog, new[] { "approval", "sign-off", "review" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("document", "file", "record"), catalog, new[] { "document", "file", "record" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("message", "email", "chat", "notification"), catalog, new[] { "message", "email", "chat", "notification" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("meeting", "workshop", "sync"), catalog, new[] { "meeting", "workshop", "sync" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("business system", "business application", "application platform"), catalog, new[] { "system", "application", "platform" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("team", "department", "business group", "workgroup"), catalog, new[] { "team", "department", "group" }, resizeToStencil, ref mapped);
            MapIfFound(builder, UnstenciledIdentity("person", "requester", "actor", "user"), catalog, new[] { "person", "user", "actor" }, resizeToStencil, ref mapped);

            if (mapped == 0) {
                throw new InvalidOperationException($"Catalog '{catalog.Name}' does not contain any collaboration or business-process-compatible stencil shapes.");
            }

            return builder;
        }

        private static void MapIfFound(VisioStencilMigrationMapBuilder builder, string masterNameU, VisioStencilCatalog catalog, string[] queries, bool resizeToStencil, ref int mapped) {
            if (catalog.TryFindBest(queries, out VisioStencilShape? stencil) && stencil != null) {
                builder.Map(UnstenciledMaster(masterNameU), stencil, resizeToStencil);
                mapped++;
            }
        }

        private static void MapIfFound(VisioStencilMigrationMapBuilder builder, Func<VisioShape, bool> predicate, VisioStencilCatalog catalog, string[] queries, bool resizeToStencil, ref int mapped) {
            if (catalog.TryFindBest(queries, out VisioStencilShape? stencil) && stencil != null) {
                builder.Map(predicate, stencil, resizeToStencil);
                mapped++;
            }
        }

        private static Func<VisioShape, bool> UnstenciledMaster(string masterNameU) {
            return shape =>
                shape != null &&
                string.IsNullOrWhiteSpace(shape.GetUserCellValue(VisioSemanticUserCells.StencilId)) &&
                string.Equals(shape.MasterNameU, masterNameU, StringComparison.OrdinalIgnoreCase);
        }

        private static Func<VisioShape, bool> UnstenciledIdentity(params string[] terms) {
            return shape => shape != null &&
                            string.IsNullOrWhiteSpace(shape.GetUserCellValue(VisioSemanticUserCells.StencilId)) &&
                            ContainsAny(IdentityText(shape), terms);
        }

        private static string IdentityText(VisioShape shape) {
            return ((shape.Text ?? string.Empty) + " " +
                    (shape.Name ?? string.Empty) + " " +
                    (shape.NameU ?? string.Empty) + " " +
                    (shape.MasterNameU ?? string.Empty)).ToLower(CultureInfo.InvariantCulture);
        }

        private static bool ContainsAny(string identity, string[] terms) {
            foreach (string term in terms) {
                if (!string.IsNullOrWhiteSpace(term) &&
                    identity.IndexOf(term.ToLower(CultureInfo.InvariantCulture), StringComparison.Ordinal) >= 0) {
                    return true;
                }
            }

            return false;
        }
    }
}
