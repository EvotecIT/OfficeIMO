using OfficeIMO.Visio.Diagrams;
using OfficeIMO.Visio.Stencils;

namespace OfficeIMO.Visio {
    public static partial class VisioGallery {
        /// <summary>
        /// Creates a data-driven hybrid network operations graph that demonstrates racks, servers, storage, edge routing, and telemetry.
        /// </summary>
        /// <param name="filePath">Target VSDX file path.</param>
        public static VisioDocument CreateHybridNetworkOperationsGraph(string filePath) {
            VisioGraphNodeRecord users = CreateNode("users", "Remote Users", VisioStencils.CollaborationBusiness, "person", "actor");
            users.IsRoot = true;
            users.ShapeData.Add("Population", "Hybrid workforce");

            VisioGraphNodeRecord internet = CreateNode("internet", "Internet", VisioStencils.Network, "internet", "wan");
            internet.ShapeData.Add("Zone", "External");

            VisioGraphNodeRecord edgeFirewall = CreateNode("edge-firewall", "Edge Firewall", VisioStencils.Network, "firewall", "security", "edge");
            edgeFirewall.ShapeData.Add("Policy", "Deny inbound by default");
            edgeFirewall.HyperlinkAddress = "https://example.org/runbooks/edge-firewall";
            edgeFirewall.HyperlinkDescription = "Edge firewall runbook";

            VisioGraphNodeRecord wanRouter = CreateNode("wan-router", "WAN Router", VisioStencils.Network, "router", "gateway");
            wanRouter.ShapeData.Add("Carrier", "Dual ISP");

            VisioGraphNodeRecord coreSwitch = CreateNode("core-switch", "Core Switch", VisioStencils.Network, "switch", "lan");
            coreSwitch.ShapeData.Add("Role", "L3 aggregation");

            VisioGraphNodeRecord rackA = CreateNode("rack-a", "Rack A", VisioStencils.Infrastructure, "rack", "cabinet", "datacenter");
            rackA.ShapeData.Add("Power", "A feed");

            VisioGraphNodeRecord rackB = CreateNode("rack-b", "Rack B", VisioStencils.Infrastructure, "rack", "cabinet", "datacenter");
            rackB.ShapeData.Add("Power", "B feed");

            VisioGraphNodeRecord hypervisorA = CreateNode("hypervisor-a", "Hypervisor A", VisioStencils.Infrastructure, "server", "host", "compute");
            hypervisorA.ShapeData.Add("Cluster", "Compute-01");

            VisioGraphNodeRecord hypervisorB = CreateNode("hypervisor-b", "Hypervisor B", VisioStencils.Infrastructure, "server", "host", "compute");
            hypervisorB.ShapeData.Add("Cluster", "Compute-01");

            VisioGraphNodeRecord storageArray = CreateNode("storage-array", "Storage Array", VisioStencils.Infrastructure, "storage-array", "san", "disk");
            storageArray.ShapeData.Add("Tier", "Replication target");

            VisioGraphNodeRecord loadBalancer = CreateNode("load-balancer", "Load Balancer", VisioStencils.Cloud, "gateway", "ingress", "load-balancer");
            loadBalancer.ShapeData.Add("Mode", "Active/active");

            VisioGraphNodeRecord appTier = CreateNode("app-tier", "App Tier", VisioStencils.Infrastructure, "server", "compute", "node");
            appTier.ShapeData.Add("Service", "Customer portal");

            VisioGraphNodeRecord database = CreateNode("database", "Database", VisioStencils.DataPlatform, "database", "sql");
            database.ShapeData.Add("Classification", "Confidential");

            VisioGraphNodeRecord backup = CreateNode("backup", "Backup NAS", VisioStencils.Network, "storage", "nas", "backup");
            backup.ShapeData.Add("Retention", "35 days");

            VisioGraphNodeRecord monitor = CreateNode("monitor", "Operations Monitor", VisioStencils.Cloud, "monitoring", "metrics", "logs");
            monitor.ShapeData.Add("Signals", "SNMP, syslog, metrics");

            VisioGraphNodeRecord noc = CreateNode("noc", "NOC Review", VisioStencils.CollaborationBusiness, "person", "actor");
            noc.ShapeData.Add("Cadence", "24x7");

            VisioGraphEdgeRecord remoteAccess = new("users", "internet") {
                Label = "VPN"
            };
            remoteAccess.ShapeData.Add("Auth", "MFA");

            VisioGraphEdgeRecord ingress = new("internet", "edge-firewall") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "ingress"
            };
            ingress.ShapeData.Add("Policy", "filtered inbound");

            VisioGraphEdgeRecord route = new("edge-firewall", "wan-router") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "route"
            };
            route.ShapeData.Add("Inspection", "IPS");

            VisioGraphEdgeRecord uplink = new("wan-router", "core-switch") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "10Gb uplink"
            };
            uplink.ShapeData.Add("Protocol", "802.1Q");

            VisioGraphEdgeRecord rackAFeed = new("core-switch", "rack-a") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "rack A"
            };

            VisioGraphEdgeRecord rackBFeed = new("core-switch", "rack-b") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "rack B"
            };

            VisioGraphEdgeRecord hostA = new("rack-a", "hypervisor-a") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "hosts"
            };
            hostA.ShapeData.Add("Role", "compute");

            VisioGraphEdgeRecord hostB = new("rack-b", "hypervisor-b") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "hosts"
            };
            hostB.ShapeData.Add("Role", "compute");

            VisioGraphEdgeRecord balance = new("load-balancer", "app-tier") {
                Kind = VisioGraphConnectorKind.Control,
                Label = "HTTPS"
            };
            balance.ShapeData.Add("Port", "443");

            VisioGraphEdgeRecord appToDb = new("app-tier", "database") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "SQL"
            };
            appToDb.ShapeData.Add("Port", "1433");

            VisioGraphEdgeRecord appToStorage = new("app-tier", "storage-array") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "block IO"
            };
            appToStorage.ShapeData.Add("Protocol", "iSCSI");

            VisioGraphEdgeRecord backupFlow = new("database", "backup") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "backup"
            };
            backupFlow.ShapeData.Add("Schedule", "nightly");

            VisioGraphEdgeRecord telemetryCore = new("core-switch", "monitor") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "telemetry"
            };
            telemetryCore.ShapeData.Add("Protocol", "SNMP");

            VisioGraphEdgeRecord telemetryHosts = new("hypervisor-a", "monitor") {
                Kind = VisioGraphConnectorKind.Data,
                Label = "metrics"
            };

            VisioGraphEdgeRecord review = new("noc", "monitor") {
                Label = "review",
                Directed = false
            };

            VisioGraphClusterRecord edgeCluster = new("edge-cluster", "Edge and WAN", new[] { "users", "internet", "edge-firewall", "wan-router" });
            edgeCluster.ShapeData.Add("Owner", "Network Security");

            VisioGraphClusterRecord datacenterCluster = new("datacenter-cluster", "Datacenter Racks and Services", new[] { "core-switch", "rack-a", "rack-b", "hypervisor-a", "hypervisor-b", "storage-array", "load-balancer", "app-tier", "database", "backup" });
            datacenterCluster.ShapeData.Add("Owner", "Infrastructure");
            datacenterCluster.HyperlinkAddress = "https://example.org/runbooks/datacenter";
            datacenterCluster.HyperlinkDescription = "Datacenter operations runbook";

            VisioGraphClusterRecord operationsCluster = new("operations-cluster", "Operations Review", new[] { "monitor", "noc" });
            operationsCluster.ShapeData.Add("Owner", "SRE");

            return VisioDocument.Create(filePath)
                .GraphDiagram("Hybrid Network Operations Graph", graph => graph
                    .Title("Hybrid Network Operations Map")
                    .Theme(VisioStyleTheme.Technical())
                    .Layout(VisioGraphLayout.Layered)
                    .Direction(VisioGraphDirection.LeftToRight)
                    .Legend()
                    .PageSize(24.0, 10.8)
                    .Margins(0.85, 0.9, 0.85, 0.85)
                    .NodeSize(1.34, 0.75)
                    .Spacing(0.92, 0.92)
                    .Import(
                        new[] { users, internet, edgeFirewall, wanRouter, coreSwitch, rackA, rackB, hypervisorA, hypervisorB, storageArray, loadBalancer, appTier, database, backup, monitor, noc },
                        new[] { remoteAccess, ingress, route, uplink, rackAFeed, rackBFeed, hostA, hostB, balance, appToDb, appToStorage, backupFlow, telemetryCore, telemetryHosts, review },
                        new[] { edgeCluster, datacenterCluster, operationsCluster }));
        }
    }
}
