# OfficeIMO Open XML layer design

OfficeIMO's Word, Excel and PowerPoint engines are built on the Open XML SDK. This document describes the layer that replaces the SDK one measured part at a time: what it owns, the contract each part keeps, and the thresholds a part must meet before an engine moves onto it. The SDK stays as the test reference throughout: it checks that output is identical, it validates, and it is the benchmark baseline.

## Thresholds

A part replaces an SDK path only when it meets all of these against the SDK path, on the published benchmark workloads, with baseline and change measured back to back:

| Kind | Time | Allocation |
| --- | --- | --- |
| Reads | at least 5× faster | at most 20% of the SDK path |
| Creates, edits and saves | at least 3× faster | at least 3× less |

Fidelity is part of the bar:

- **Reads:** return the same text and structure as the SDK path, or the same as the OfficeIMO model where the model defines the semantics.
- **Saves:**
  - Untouched parts are byte for byte the same.
  - Edited parts are semantically identical XML outside the edit.
  - The SDK validator passes.
  - Microsoft Office opens the file.
- **No hidden SDK:** a path that claims to be SDK-free doesn't load the SDK assembly at run time, and a NativeAOT app that uses only that path has no SDK code in its image.

## Ownership

| Capability | Owner |
| --- | --- |
| ZIP container read and write, exact Deflate framing, CRC-32 | `OfficeIMO.Core` (`Internal/OpenXml`, `Internal/Packaging`, `Internal/Compression`) |
| OPC: part names and identity, content types, relationships, core properties | `OfficeIMO.Core` (`OfficeOpcPackage`, `OfficeOpcPartNames`) |
| XML pull reader, namespace table, markup compatibility | `OfficeIMO.Core` (`OfficeXmlPullReader`, `OfficeXmlNamespaces`) |
| Editable element tree, UTF-8 writer, package writer | `OfficeIMO.Core` |
| Typed views over the tree | Generated from the ECMA-376 schemas, into the engine that uses them |
| Package security inspection | `OfficePackageSecurityInspector`, sharing the part-name rules above |
| Validation | The Open XML SDK, as an opt-in and as the test reference |

Engines call the layer; they don't parse ZIP or XML themselves. Code that behaves differently on errors keeps its own policy until its owner moves: the Excel raw reader falls back to the SDK, and `WordZipPackage` and the provenance ZIP parser accept inputs the strict reader rejects.

## Read layer

- **ZIP.** Members come from the `OfficeArchiveSafety` central-directory scan. Each member inflates into an exact-size buffer and must end at its declared length, match its CRC-32, and have Deflate framing that ends exactly at the compressed length. Member names that differ only in case name the same part and are rejected.
- **OPC.** Part names compare by one identity: percent-decoded once, ASCII case-insensitive. Relationship targets resolve on first use. Strict relationship types read as transitional.
- **Pull reader.** It reads the UTF-8 bytes of a part (UTF-16 parts are transcoded first) and resolves prefixes per element scope. It rejects document type declarations, undefined entities, malformed UTF-8, characters XML does not allow, a missing or second root, and repeated attributes. Malformed XML throws `XmlException`.
- **Namespaces.** Each declared namespace has two ids:
  - a folded id, under which Strict and transitional URIs are one namespace, used for vocabulary;
  - an exact URI id, for XML identity and round trips.
- **Markup compatibility.** It follows ECMA-376 Part 3, with `mc:MustUnderstand` enforced on the branch that is kept. Readers that must reproduce a model's own treatment of compatibility markup read it as written instead.

Readers built on this layer follow the semantics of the model they replace. The streamed Word text projection uses the Word model's text rules; the PowerPoint slide reader uses the presentation model's shape rules. Their output is compared with the model on the repository's documents.

## Document model

### Element tree

Each parsed part is a generic tree stored in arrays: node records indexed by integer, not one object per element. A node keeps:

- its element name and its namespace, as both the folded and the exact id;
- its attributes in source order, and its namespace declarations where they were declared;
- the byte ranges of its source markup.

Typed access comes from views generated from the ECMA-376 schemas, for the types an engine uses. A view is a small value over a tree and a node index. It reads and writes attributes with schema value types, and inserts children in schema order.

### Laziness and passthrough

A part stays as its compressed ZIP member until something reads or edits it.

- **On save:** an untouched member is copied as stored, compressed bytes included. A parsed part whose subtree is unchanged writes its source bytes back verbatim. Only edited elements are serialised.
- **Large parts:** parsing only the depth an edit needs, and keeping unvisited subtrees as source ranges, is the next step for parts like Word's `document.xml`. It needs the namespace scope at each range boundary, which the reader already tracks.

### Unknown markup

What was read is kept as it was read:

- prefixes and where namespaces were declared;
- attribute order;
- unknown elements and attributes, including ignorable ones;
- `mc:AlternateContent` as written;
- comments, processing instructions and whitespace.

New elements use the prefix already in scope for their namespace, with its exact URI. A Strict part therefore gets Strict names for new content. Where no prefix is in scope, the canonical prefix is declared where the SDK would declare it.

### Write path

Edited parts are serialised as UTF-8 straight into the Deflate encoder, with no intermediate strings or DOM. The ZIP writer and Deflate encoder that Excel's direct writer uses are the shared writer.

### Public API

About 40 public members expose SDK types.

- **Escape hatches** (raw access to parts and elements) move to an `OfficeIMO.*.OpenXmlSdk` interop package of extension methods, so the core assemblies don't reference the SDK.
- **Other members**, which expose SDK types only for convenience, get OfficeIMO types in the major version that removes the SDK, with a migration note.

### Validation

`OpenXmlValidator` remains the validator, exposed as an opt-in through the interop package. Every model test saves through the layer's writer and validates the output with the SDK.

## Open decisions

- **Reader model fallback.** The Reader's Word and PowerPoint handlers stream text without the SDK and fall back to the model when a document uses something the streamed pass doesn't read. The fallback stays built in. Whether it becomes opt-in is undecided. That decision sets the NativeAOT size of apps built on the Reader handlers: the SDK stays in the image while the fallback is reachable.

## Order of work

The read layer comes first, then a PowerPoint prototype of the model. The prototype has to meet the thresholds on these workloads:

- 200-slide create, edit and save against the SDK;
- a lossless open-and-save of the PowerPoint corpus;
- the validator passing;
- no SDK in the NativeAOT image of an app that uses only the model.

Engines move only after that, one feature per change, PowerPoint first. The SDK leaves the runtime dependencies in a major version.
