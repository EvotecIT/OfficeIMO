namespace OfficeIMO.Access {
    /// <summary>One operation's qualified boundary. A detected extension or profile does not imply support.</summary>
    public sealed class AccessOperationCapability {
        internal AccessOperationCapability(string operation, bool supported, string boundary) { Operation = operation; IsSupported = supported; Boundary = boundary; }
        /// <summary>Stable operation identifier.</summary>
        public string Operation { get; }
        /// <summary>Whether the described operation has implemented qualification.</summary>
        public bool IsSupported { get; }
        /// <summary>Exact evidence and limitation boundary.</summary>
        public string Boundary { get; }
    }

    /// <summary>Canonical operation catalog used by consumers and the generated support matrix.</summary>
    public static class AccessCapabilities {
        /// <summary>Current capability contracts. Native header recognition remains distinct from catalog decoding.</summary>
        public static IReadOnlyList<AccessOperationCapability> Operations { get; } = Array.AsReadOnly(new[] {
            new AccessOperationCapability("model.create", true, "New Jet4/ACE12-targeted in-memory model with separately assessed native creation."),
            new AccessOperationCapability("model.edit", true, "Typed tables, columns, primary-key definitions, relationships, inert query text and rows; rollback and read leases."),
            new AccessOperationCapability("model.rows.read", true, "Forward-only DbDataReader over modeled or qualified native tables; modeled input retains omitted/null distinctions."),
            new AccessOperationCapability("native.header.inspect", true, "Inert bounded header/profile, page-alignment, byte/page limits and snapshot SHA-256. Independent Jet4/ACE12/ACE14/ACE16/ACE17 fixtures; Jet3 recognized without producer qualification. Protection and feature values are not decoded."),
            new AccessOperationCapability("native.catalog.read", true, "Unprotected Jet4 and ACE12/14/16/17 catalog, selected table/column properties, index definitions, relationships and inert links. System/unknown catalog entries remain separate."),
            new AccessOperationCapability("native.rows.read", true, "Bounded forward-only native scans, deleted/overflow/fragmented rows, lazy fields and incremental binary streams. Qualified scalar/Unicode/Memo/OLE, complex values, attachments, BigInt and extended-date profiles; calculated/unknown values retain bytes and diagnostics."),
            new AccessOperationCapability("native.query.records.read", true, "Exact inert query records and typed parameter inventory. Simple single-table SELECT and two-part UNION SQL reconstruction is qualified; other query SQL remains explicitly unavailable."),
            new AccessOperationCapability("native.create", true, "Seed-free unprotected Jet4/ACE12 output from new table models; qualified scalar/Unicode/long values, sequential AutoNumber seeds, primary/unique/composite indexes and enforced single-field relationships. General legacy ASCII collation subset; unsupported types, keys, queries and lossy values fail before output."),
            new AccessOperationCapability("native.preserve", true, "Immutable native snapshot copied without byte changes to the same family/profile, including opaque/protected/compiled content. Source identity is checked; signature validity and protected-content decoding are not assessed."),
            new AccessOperationCapability("native.edit", false, "Existing native documents remain immutable; row/schema/application edits require separate codecs."),
            new AccessOperationCapability("native.convert", false, "Known complex-field, rich-text, BigInt, extended-date and unqualified opaque-property target mappings are diagnosed. Native conversion and persistence remain unavailable."),
            new AccessOperationCapability("application.objects.read", true, "Jet4/ACE application inventory and exact payloads; qualified ACE version-21 designer sections/controls/sources/event metadata, 76-byte StopMacro definitions, table data-macro XML and resource metadata. Jet version-19 designers and unknown definitions are preserve-only. Dependency references are partial and inert."),
            new AccessOperationCapability("application.objects.write", false, "No native form/report/macro/module writer is qualified."),
            new AccessOperationCapability("vba.inspect", true, "Access Jet compound storage and ACE hierarchy adapt to shared Core MS-OVBA parsing. Module source and stored references are decoded when available; compiled/unknown streams remain exact. Built-in references and signature validation are outside this contract."),
            new AccessOperationCapability("protection.inspect", false, "Password, encryption and signature state is NotAssessed."),
            new AccessOperationCapability("report.render", false, "Native report definitions and expression evaluation are not qualified.")
        });
    }
}
