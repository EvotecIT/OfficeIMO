using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Visio.Stencils {
    /// <summary>
    /// Built-in OfficeIMO-native stencil catalogs.
    /// </summary>
    public static class VisioStencils {
        /// <summary>
        /// Gets basic geometry shapes.
        /// </summary>
        public static VisioStencilCatalog BasicShapes { get; } = new(
            "Basic Shapes",
            new[] {
                Shape("basic.rectangle", "Rectangle", "Rectangle", "Basic Shapes", 2.0, 1.0, "box", "process-box"),
                Shape("basic.square", "Square", "Square", "Basic Shapes", 1.2, 1.2, "box"),
                Shape("basic.circle", "Circle", "Circle", "Basic Shapes", 1.2, 1.2, "round"),
                Shape("basic.ellipse", "Ellipse", "Ellipse", "Basic Shapes", 2.0, 1.0, "oval"),
                Shape("basic.diamond", "Diamond", "Diamond", "Basic Shapes", 1.6, 1.2, "decision"),
                Shape("basic.triangle", "Triangle", "Triangle", "Basic Shapes", 1.6, 1.3),
                Shape("basic.parallelogram", "Parallelogram", "Parallelogram", "Basic Shapes", 2.0, 1.0, "data"),
                Shape("basic.hexagon", "Hexagon", "Hexagon", "Basic Shapes", 2.0, 1.0, "preparation"),
                Shape("basic.trapezoid", "Trapezoid", "Trapezoid", "Basic Shapes", 2.0, 1.0),
                Shape("basic.pentagon", "Pentagon", "Pentagon", "Basic Shapes", 1.5, 1.3)
            });

        /// <summary>
        /// Gets common flowchart shapes.
        /// </summary>
        public static VisioStencilCatalog Flowchart { get; } = new(
            "Flowchart",
            new[] {
                Shape("flow.process", "Process", "Process", "Flowchart", 2.4, 1.0, "step", "task"),
                Shape("flow.decision", "Decision", "Decision", "Flowchart", 2.0, 1.4, "branch", "choice"),
                Shape("flow.data", "Data", "Data", "Flowchart", 2.4, 1.0, "input", "output"),
                Shape("flow.preparation", "Preparation", "Preparation", "Flowchart", 2.2, 1.0, "setup"),
                Shape("flow.manual-operation", "Manual Operation", "Manual operation", "Flowchart", 2.4, 1.0, "manual"),
                Shape("flow.off-page-reference", "Off-page Reference", "Off-page reference", "Flowchart", 1.0, 1.0, "continuation", "off-page")
            });

        /// <summary>
        /// Gets block diagram building blocks.
        /// </summary>
        public static VisioStencilCatalog BlockDiagram { get; } = new(
            "Block Diagram",
            new[] {
                Shape("block.block", "Block", "Rectangle", "Block Diagram", 2.4, 1.0, "module", "component"),
                Shape("block.storage", "Storage", "Data", "Block Diagram", 2.4, 1.0, "data-store"),
                Shape("block.decision", "Decision Block", "Decision", "Block Diagram", 2.0, 1.4, "branch"),
                Shape("block.region", "Region", "Rectangle", "Block Diagram", 3.5, 2.0, "container", "group")
            });

        /// <summary>
        /// Gets dependency-free architecture and infrastructure shapes.
        /// </summary>
        public static VisioStencilCatalog Architecture { get; } = new(
            "Architecture",
            new[] {
                Shape("arch.actor", "Actor", "Circle", "Architecture", 0.9, 0.9, "user", "person", "client"),
                Shape("arch.service", "Service", "Process", "Architecture", 1.8, 1.0, "app", "application", "api"),
                Shape("arch.compute", "Compute", "Process", "Architecture", 1.8, 1.0, "vm", "server", "worker", "container"),
                Shape("arch.gateway", "Gateway", "Decision", "Architecture", 1.7, 1.1, "ingress", "load-balancer", "endpoint"),
                Shape("arch.database", "Database", "Data", "Architecture", 1.8, 1.0, "sql", "data-store"),
                Shape("arch.storage", "Storage", "Data", "Architecture", 1.8, 1.0, "blob", "file", "object-store"),
                Shape("arch.queue", "Queue", "Data", "Architecture", 1.8, 1.0, "bus", "stream", "broker"),
                Shape("arch.security", "Security", "Decision", "Architecture", 1.7, 1.1, "identity", "key", "policy"),
                Shape("arch.network", "Network", "Rectangle", "Architecture", 2.2, 1.0, "subnet", "vnet", "route"),
                Shape("arch.region", "Region", "Rectangle", "Architecture", 4.0, 2.4, "container", "boundary")
            });

        /// <summary>
        /// Gets dependency-free network and infrastructure shapes.
        /// </summary>
        public static VisioStencilCatalog Network { get; } = new(
            "Network",
            new[] {
                Shape("net.user", "User", "Circle", "Network", 0.9, 0.9, "client", "person"),
                Shape("net.workstation", "Workstation", "Process", "Network", 1.45, 0.85, "desktop", "laptop", "endpoint"),
                Shape("net.server", "Server", "Process", "Network", 1.45, 0.85, "vm", "host"),
                Shape("net.switch", "Switch", "Rectangle", "Network", 1.65, 0.65, "hub", "lan"),
                Shape("net.router", "Router", "Decision", "Network", 1.35, 1.0, "route", "gateway"),
                Shape("net.firewall", "Firewall", "Decision", "Network", 1.35, 1.0, "security", "edge"),
                Shape("net.internet", "Internet", "Circle", "Network", 0.95, 0.95, "wan", "cloud"),
                Shape("net.printer", "Printer", "Process", "Network", 1.45, 0.85, "peripheral"),
                Shape("net.storage", "Storage", "Data", "Network", 1.45, 0.85, "nas", "san"),
                Shape("net.database", "Database", "Data", "Network", 1.45, 0.85, "sql"),
                Shape("net.wireless", "Wireless AP", "Circle", "Network", 0.95, 0.95, "wifi", "access-point"),
                Shape("net.zone", "Zone", "Rectangle", "Network", 4.0, 2.4, "container", "boundary")
            });

        /// <summary>
        /// Gets cross-functional swimlane and process-map shapes.
        /// </summary>
        public static VisioStencilCatalog Swimlane { get; } = new(
            "Swimlane",
            new[] {
                Shape("swim.activity", "Activity", "Process", "Swimlane", 1.6, 0.72, "step", "task", "process"),
                Shape("swim.decision", "Decision", "Decision", "Swimlane", 1.45, 0.95, "branch", "choice"),
                Shape("swim.data", "Data", "Data", "Swimlane", 1.6, 0.72, "input", "output", "document"),
                Shape("swim.start-end", "Start/End", "Ellipse", "Swimlane", 1.5, 0.72, "terminator", "start", "end"),
                Shape("swim.lane", "Lane", "Rectangle", "Swimlane", 6.0, 1.45, "role", "participant", "container"),
                Shape("swim.phase", "Phase", "Rectangle", "Swimlane", 2.4, 0.55, "milestone", "stage", "column")
            });

        /// <summary>
        /// Gets organization chart shapes.
        /// </summary>
        public static VisioStencilCatalog OrgChart { get; } = new(
            "Org Chart",
            new[] {
                Shape("org.executive", "Executive", "Process", "Org Chart", 1.85, 0.82, "root", "leader", "ceo"),
                Shape("org.manager", "Manager", "Process", "Org Chart", 1.85, 0.82, "lead", "supervisor"),
                Shape("org.position", "Position", "Process", "Org Chart", 1.85, 0.82, "person", "employee", "role"),
                Shape("org.assistant", "Assistant", "Rectangle", "Org Chart", 1.7, 0.65, "ea", "staff"),
                Shape("org.vacancy", "Vacancy", "Rectangle", "Org Chart", 1.85, 0.82, "open", "hiring"),
                Shape("org.external", "External", "Rectangle", "Org Chart", 1.85, 0.82, "advisor", "vendor", "partner"),
                Shape("org.team-band", "Team Band", "Rectangle", "Org Chart", 4.0, 1.6, "department", "container", "group")
            });

        /// <summary>
        /// Gets timeline and roadmap shapes.
        /// </summary>
        public static VisioStencilCatalog Timeline { get; } = new(
            "Timeline",
            new[] {
                Shape("time.axis", "Timeline Axis", "Rectangle", "Timeline", 8.0, 0.06, "roadmap", "schedule"),
                Shape("time.milestone", "Milestone", "Diamond", "Timeline", 0.25, 0.25, "date", "marker"),
                Shape("time.release", "Release", "Circle", "Timeline", 0.25, 0.25, "delivery", "launch"),
                Shape("time.decision", "Decision", "Circle", "Timeline", 0.25, 0.25, "approval", "gate"),
                Shape("time.risk", "Risk", "Circle", "Timeline", 0.25, 0.25, "issue", "attention"),
                Shape("time.span", "Span", "Rectangle", "Timeline", 2.0, 0.28, "phase", "workstream", "duration"),
                Shape("time.label", "Label", "Rectangle", "Timeline", 1.45, 0.48, "annotation", "callout")
            });

        /// <summary>
        /// Gets UML-style sequence diagram shapes.
        /// </summary>
        public static VisioStencilCatalog Sequence { get; } = new(
            "Sequence Diagram",
            new[] {
                Shape("seq.participant", "Participant", "Rectangle", "Sequence Diagram", 1.45, 0.62, "lifeline", "service", "component"),
                Shape("seq.actor", "Actor", "Circle", "Sequence Diagram", 0.72, 0.72, "user", "person", "client"),
                Shape("seq.boundary", "Boundary", "Rectangle", "Sequence Diagram", 1.45, 0.62, "edge", "interface"),
                Shape("seq.control", "Control", "Rectangle", "Sequence Diagram", 1.45, 0.62, "coordinator", "controller"),
                Shape("seq.entity", "Entity", "Rectangle", "Sequence Diagram", 1.45, 0.62, "domain", "object"),
                Shape("seq.database", "Database", "Data", "Sequence Diagram", 1.45, 0.62, "store", "data-store"),
                Shape("seq.note", "Note", "Rectangle", "Sequence Diagram", 1.8, 0.75, "annotation", "callout")
            });

        /// <summary>
        /// Gets a combined catalog containing all built-in OfficeIMO-native stencil shapes.
        /// </summary>
        public static VisioStencilCatalog All { get; } = new(
            "All Built-in Stencils",
            BasicShapes.Shapes
                .Concat(Flowchart.Shapes)
                .Concat(BlockDiagram.Shapes)
                .Concat(Architecture.Shapes)
                .Concat(Network.Shapes)
                .Concat(Swimlane.Shapes)
                .Concat(OrgChart.Shapes)
                .Concat(Timeline.Shapes)
                .Concat(Sequence.Shapes)
                .GroupBy(shape => shape.Id)
                .Select(group => group.First())
                .ToList());

        private static VisioStencilShape Shape(
            string id,
            string name,
            string masterNameU,
            string category,
            double defaultWidth,
            double defaultHeight,
            params string[] keywords) {
            string prefix = id.Contains(".") ? id.Substring(0, id.IndexOf('.')) : id;
            string localId = id.Contains(".") ? id.Substring(id.IndexOf('.') + 1) : id;
            string[] tags = new[] { prefix, category, masterNameU };
            string[] aliases = keywords
                .Concat(new[] { localId, name.Replace(" ", "-") })
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToArray();
            return new VisioStencilShape(id, name, masterNameU, category, defaultWidth, defaultHeight, keywords, aliases, tags, masterNameU, VisioMeasurementUnit.Inches);
        }
    }
}
