using System.Collections.Generic;

namespace OfficeIMO.DocBook;

/// <summary>Tracks bounded common-profile sibling constraints in one forward traversal.</summary>
internal sealed class DocBookSiblingState {
    private HashSet<DocBookNodeKind>? _singletons;
    internal bool SawInfo { get; private set; }
    internal bool SawNonInfo { get; private set; }
    internal bool SawBody { get; private set; }
    internal bool SawSubdivision { get; private set; }
    internal bool SawTableFoot { get; private set; }
    internal int MaximumCalsOrder { get; private set; } = -1;

    internal bool HasSingleton(DocBookNodeKind kind) => _singletons?.Contains(kind) == true;

    internal void Record(DocBookNodeKind kind, bool isComponentInfo, bool isTableFoot, int calsOrder, bool isSingleton) {
        SawInfo |= isComponentInfo;
        SawNonInfo |= kind != DocBookNodeKind.Info;
        SawBody |= kind is not (DocBookNodeKind.Info or DocBookNodeKind.Title or DocBookNodeKind.Subtitle);
        SawSubdivision |= kind is DocBookNodeKind.Section or DocBookNodeKind.Index;
        SawTableFoot |= isTableFoot;
        if (calsOrder > MaximumCalsOrder) MaximumCalsOrder = calsOrder;
        if (isSingleton) (_singletons ??= new HashSet<DocBookNodeKind>()).Add(kind);
    }
}
