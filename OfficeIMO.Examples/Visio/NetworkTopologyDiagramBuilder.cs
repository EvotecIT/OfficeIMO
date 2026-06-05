using System;
using System.IO;
using OfficeIMO.Visio;
using OfficeIMO.Visio.Diagrams;

namespace OfficeIMO.Examples.Visio {
    public static class NetworkTopologyDiagramBuilder {
        public static void Example_NetworkTopologyDiagramBuilder(string folderPath, bool openVisio) {
            Console.WriteLine("[*] Visio - Network topology diagram builder");
            string filePath = Path.Combine(folderPath, "Network Topology Diagram Builder.vsdx");

            VisioDocument.Create(filePath)
                .NetworkTopologyDiagram("Branch Topology", topology => topology
                    .PageSize(12.5, 9)
                    .Theme(VisioStyleTheme.Technical())
                    .Spacing(1.35, 1.1)
                    .Root("internet", "Internet", VisioNetworkNodeKind.Internet)
                    .Firewall("firewall", "Firewall")
                    .Switch("core", "Core Switch")
                    .Server("app", "App Server")
                    .Server("auth", "Auth Proxy")
                    .Database("db", "Database")
                    .Storage("backup", "Backup NAS")
                    .Workstation("finance", "Finance PC")
                    .Workstation("support", "Support PC")
                    .Printer("printer", "Printer")
                    .Wireless("wifi", "Wi-Fi")
                    .Subnet("edge", "Edge", "internet", "firewall", "core")
                    .Subnet("servers", "Server Zone", "app", "db", "backup")
                    .Subnet("clients", "Client LAN", "finance", "support", "printer", "wifi")
                    .Ethernet("internet", "firewall", "WAN")
                    .Trunk("firewall", "core", "uplink")
                    .Trunk("core", "app", "10Gb")
                    .Trunk("core", "auth", "TLS")
                    .Ethernet("app", "db")
                    .Management("auth", "backup", "keys")
                    .Management("app", "backup", "backup")
                    .Ethernet("core", "finance")
                    .Ethernet("core", "support")
                    .Ethernet("support", "printer")
                    .WirelessLink("core", "wifi", "wireless"))
                .Save();

            if (openVisio) {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo(filePath) { UseShellExecute = true });
            }
        }
    }
}
