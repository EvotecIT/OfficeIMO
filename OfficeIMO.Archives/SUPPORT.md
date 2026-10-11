# Archive support

| Container | Supported read profile | Boundaries |
| --- | --- | --- |
| CAB | Single cabinets with Store, MSZIP and the shared LZX decoder | Cabinet sets, continuation/spanning and unsupported compression are rejected; finite entry/folder/output bounds and available checksums are enforced |
| MSI | Ordinary installer database CFB; schema-driven strings and Directory, Component, File and Media tables; embedded complete cabinets | External cabinets and loose files, transforms, patches and action-dependent installation layouts are unsupported; no Installer actions, conditions or properties are executed |

CAB member names support ASCII or explicitly flagged UTF-8, with at most 256 encoded bytes. Unflagged names containing non-ASCII bytes fail explicitly because the container does not identify the producer's locale. Unsupported CAB compression, continuation and filename-encoding profiles raise `NotSupportedException`; corrupt or unsafe content raises `InvalidDataException`.

MSI strings support two- and three-byte references, extended lengths, nullable 16-/32-bit integers and authored short/long names. Metadata schemas determine field widths: older and newer packages can use different `Media.LastSequence` widths. Unsupported code pages fail explicitly. Code page zero accepts ASCII deterministically rather than inheriting a host locale.

On .NET 8/10, legacy string decoding uses the runtime code-page provider directly without registering it globally. The portable netstandard2.0 assembly includes OfficeIMO.Core's supported single-byte fallback encodings (1250, 1252, 437, 850 and Mac Roman); additional encodings depend on the host's available strict .NET decoder. UTF-8 is supported on every target. .NET Framework supplies its installed code-page support.

MSI CAB member keys are mapped to authored File-table names and directory paths. Directory identifiers, references and embedded cabinet stream names preserve their case-sensitive MSI identity; File identifiers match cabinet members without case sensitivity. Media sequence ranges follow `DiskId` order. Filename case is preserved. The extractor does not reproduce an installation: feature selection, registry state, transforms, custom actions and environment folder properties do not alter the extracted payload.

MSI directory depth counts the root and directories authored as `.`, including directories with no payload files. Cached ancestors retain their complete depth. Output segment and path-length bounds are checked before a resolved directory prefix is assembled or cached.

MSI compound storage is indexed once per extraction. Embedded cabinet input bytes, members and decoded bytes share aggregate limits. `MaxCabinetCount` defaults to 256 and includes cabinets whose mapped files are empty or excluded by `IncludeFile`.

Names are checked for absolute paths, traversal, alternate data streams, reserved Windows device names, invalid portable segments, case/Unicode collisions and file/directory conflicts. The destination, staging directory and paths within staging must not be symbolic links or reparse points. Caller-selected parent directories and read-only input paths may traverse operating-system or caller-owned links. Existing files are not overwritten. Atomic publication uses a same-parent directory move; callers must coordinate access to the destination and its parent during extraction.

Focused qualification covers Windows administrative-template MSI payload profiles and Microsoft Edge CAB payloads. It does not imply compatibility with every MSI or LZX cabinet producer. Runtime package availability and source qualification are separate release states.
