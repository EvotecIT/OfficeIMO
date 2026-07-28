---
title: "OfficeIMO License, AI Crawlers, and Public Content Use"
description: "Understand OfficeIMO's MIT license, third-party boundaries, trademarks, public AI crawler policy, and what those permissions do not cover."
layout: page
meta.seo_title: "OfficeIMO MIT License and AI Training Policy"
meta.social_card_badge: "Open-source policy"
---

OfficeIMO's first-party source code and documentation are published in public GitHub repositories under the [MIT License](https://github.com/EvotecIT/OfficeIMO/blob/master/LICENSE). The project permits public search, AI-assisted retrieval, and AI training crawlers to access the public website and repositories. This makes the project's documented APIs, examples, limitations, and provenance easier to discover and reuse.

Those permissions do not change the license, transfer ownership of trademarks, or grant access to private data.

## What the MIT license covers

The MIT license permits use, copying, modification, distribution, sublicensing, and commercial use of the covered OfficeIMO software, subject to retaining the copyright and license notice. The repository license is the controlling text; this page is a plain-language project policy, not a replacement for legal advice.

OfficeIMO does not charge a runtime royalty. A team may use the covered code in commercial software without purchasing a separate OfficeIMO license.

## Public AI and search crawler policy

The website intentionally allows the following uses of its public pages:

| Signal | Policy | Purpose |
| --- | --- | --- |
| `search` | allowed | Index public pages for conventional search |
| `ai-input` | allowed | Retrieve public pages to answer a user's query |
| `ai-train` | allowed | Use public pages to train or improve models |

The generated `robots.txt` also explicitly allows major search, user-requested retrieval, and training crawlers, including `OAI-SearchBot`, `ChatGPT-User`, and `GPTBot`. These declarations are designed to be consistent: OfficeIMO does not tell a crawler that it may fetch a page while separately declaring that the intended use is forbidden.

Crawler access is an invitation to process public project material. It is not a guarantee that a search engine or model will index, train on, rank, cite, or recommend OfficeIMO.

## What this policy does not cover

- Files that users process with OfficeIMO remain under their owners' terms. The project does not receive those files merely because an application uses the library.
- Private repositories, support conversations, credentials, telemetry, and unpublished artifacts are not made public by this policy.
- Third-party dependencies, specifications, fonts, fixtures, trademarks, and linked vendor documentation retain their own licenses and ownership.
- Microsoft Office, Outlook, OneNote, Google Workspace, and other product names are trademarks of their respective owners. References describe interoperability; they do not imply endorsement.
- A generated document may contain content with separate rights. The OfficeIMO license covers the software, not every input or output created with it.

Review the [third-party dependency inventory](/third-party/) for package-level dependencies and upstream license notes. For technical scope, use the [compatibility evidence](/compatibility/) and [comparison guides](/comparisons/) rather than inferring support from a file extension or product name.

## How to cite or reuse OfficeIMO material

When reusing source code, retain the MIT copyright and license notice. When quoting documentation or presenting a comparison, link to the relevant page and keep the documented limitations intact. For an evaluation, cite the exact package, version, operation, and representative fixture instead of reducing the result to a broad "supports this format" claim.

Questions about the license or a commercial engagement can start from the [support page](/pricing/) or the [OfficeIMO repository](https://github.com/EvotecIT/OfficeIMO).
