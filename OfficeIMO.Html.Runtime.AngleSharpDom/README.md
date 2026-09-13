# Retained AngleSharp DOM provider

This internal project builds the runtime worker's AngleSharp dependency from
pinned source so mutation notifications can join the interpreter's microtask
queue. It remains a retained parser and DOM dependency. Ordinary OfficeIMO HTML
consumers continue to use the original AngleSharp NuGet package.

`Upstream/` retains the C# sources from
[AngleSharp 1.7.1 commit fbfd8db](https://github.com/AngleSharp/AngleSharp/tree/fbfd8dbf5670ad77a986ebf128377693ecfa4e08/src/AngleSharp).
`upstream.json` records the original paths and Git blob hashes. The original MIT
license is retained in `AngleSharp.LICENSE.txt` and copied to build and publish
outputs.

The local changes have these responsibilities:

- `MutationHost` accepts an optional `IMutationMicrotaskScheduler`, maintains
  pending observer order and delivers one compound notification. Documents in
  the same agent share that queue.
- `DocumentExtensions` reports interested observers when records are queued and
  exposes `IDomMutationListener` for synchronous internal invalidation. The worker
  uses the latter for base URLs without creating a web-visible observer job.
- `MutationObserver` retains Document targets, source-associated transients and
  registration order, and clears registrations on disconnect. `Node` queues Document child changes, preserves replacement records
  and transient subtree observations, and compares actual tree roots.
- `Document`, `HtmlDocument` and `DomImplementation` let secondary HTML documents
  inherit the creating agent's mutation host while retaining an inert context;
  HTML document clones retain that host as well.
- `DocumentPositions` names `Node` as the owner of document-position constants.

`IMutationMicrotaskScheduler` and `IDomMutationListener` are narrow integration
hooks outside the retained source directory. Scheduling, script error handling,
resource authority and session lifetime belong to the worker. A future provider
can replace this assembly at those boundaries; it must pass the runtime mutation,
history and application contracts before replacing it.

The assembly keeps AngleSharp's name and assembly version for the retained CSS
and JavaScript bindings. It uses public signing with `AngleSharp.PublicKey.bin`;
no private signing key is included. This is not an upstream-signed release.
The informational version and `OfficeIMO.ProviderPatch` assembly metadata identify
the modified build, and runtime captures include its informational version.
Deploy the complete worker directory. Do not substitute this DLL into an ordinary
consumer's dependency directory. This project is not a public NuGet package.

When refreshing source, compare each file against the recorded upstream hashes,
review the local delta, retain attribution and verify complete deployed worker
outputs on supported runtimes. Keep the original consumer package and worker
provider isolation covered by deployment evidence.
