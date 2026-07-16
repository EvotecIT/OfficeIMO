namespace OfficeIMO.Excel.Xlsb.Package {
    internal sealed class XlsbPackageRelationship {
        internal XlsbPackageRelationship(string id, string type, string target, bool isExternal) {
            Id = id;
            Type = type;
            Target = target;
            IsExternal = isExternal;
        }

        internal string Id { get; }

        internal string Type { get; }

        internal string Target { get; }

        internal bool IsExternal { get; }
    }
}
