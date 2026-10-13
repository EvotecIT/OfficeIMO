---
title: "OfficeIMO for AI agents"
description: "Add OfficeIMO's MCP server, plugin and Agent Skills to Claude, Codex, GitHub Copilot, Cursor and other agents to inspect, search and convert local documents and mail."
meta.seo_title: "OfficeIMO for AI agents: MCP server, plugin and skills"
layout: application
meta.application: "agents"
---

## Documents stay on your computer

The OfficeIMO MCP server runs locally, started by your agent with `dotnet dnx`. It opens only files inside the folders you name in `OFFICEIMO_MCP_ALLOWED_ROOTS` and refuses to start without one. Nothing is uploaded by OfficeIMO; what the agent sends to its model is up to the agent and the parts it fetches.

## Search first, then fetch what matters

Agents start with a bounded inspection and a narrow search, then fetch only the results they need, a page at a time. A mailbox of thousands of messages never lands in the conversation whole, and document text is treated as untrusted content rather than instructions.

## Read freely, write only where asked

Inspection, search and fetching never change a file. Conversion and the PDF tools write a new file at a path you name, and keep existing files unless you ask to replace them. Protected PDFs keep their permissions, and passwords come from the host's environment rather than the conversation.

| Task | Ask your agent | Tools it uses |
|---|---|---|
| Summarize a report | "Summarize the risks section of Q3-report.docx" | Inspect, search, fetch |
| Find a message | "Find the email from Anna about the invoice in archive.pst" | Mail search, fetch |
| Extract for a model | "Convert handbook.pdf to Markdown in notes/" | Convert |
| Prepare PDFs | "Combine these scans into one PDF" | Assemble, and OCR where the host registered a text-recognition provider |
| Check a change | "Compare pages 1 to 5 of the old and new contract" | Compare |

The same operations are available from the [OfficeIMO command line](/tool/) for scripts and CI.
