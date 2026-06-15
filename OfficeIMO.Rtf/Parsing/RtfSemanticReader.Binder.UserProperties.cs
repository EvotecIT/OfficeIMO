using OfficeIMO.Rtf.Syntax;

namespace OfficeIMO.Rtf;

internal static partial class RtfSemanticReader {
    private sealed partial class Binder {
        private static IReadOnlyList<RtfUserProperty> ReadUserProperties(RtfGroup root, int ansiCodePage, int unicodeSkipCount) {
            RtfGroup? userPropertiesGroup = root.Children.OfType<RtfGroup>().FirstOrDefault(group => group.Destination == "userprops");
            if (userPropertiesGroup == null) return Array.Empty<RtfUserProperty>();

            var properties = new List<RtfUserProperty>();
            PendingUserProperty? pending = null;

            foreach (RtfNode node in userPropertiesGroup.Children) {
                if (node is RtfGroup childGroup) {
                    switch (childGroup.Destination) {
                        case "propname":
                            AddPendingUserProperty(properties, pending);
                            string name = CollectPlainText(childGroup, ansiCodePage, unicodeSkipCount).Trim();
                            pending = string.IsNullOrEmpty(name) ? null : new PendingUserProperty(name);
                            break;
                        case "staticval":
                            if (pending != null) {
                                pending.StaticValue = CollectPlainText(childGroup, ansiCodePage, unicodeSkipCount).Trim();
                            }
                            break;
                        case "linkval":
                            if (pending != null) {
                                pending.LinkedValue = CollectPlainText(childGroup, ansiCodePage, unicodeSkipCount).Trim();
                            }
                            break;
                    }
                } else if (node is RtfControlWord control && control.Name == "proptype" && pending != null) {
                    pending.TypeCode = control.Parameter;
                }
            }

            AddPendingUserProperty(properties, pending);
            return properties;
        }

        private static void AddPendingUserProperty(List<RtfUserProperty> properties, PendingUserProperty? pending) {
            if (pending == null) return;
            properties.Add(new RtfUserProperty(pending.Name, pending.TypeCode, pending.StaticValue) {
                LinkedValue = pending.LinkedValue
            });
        }

        private sealed class PendingUserProperty {
            public PendingUserProperty(string name) {
                Name = name;
            }

            public string Name { get; }

            public int? TypeCode { get; set; }

            public string? StaticValue { get; set; }

            public string? LinkedValue { get; set; }
        }
    }
}
