# OfficeIMO SEO and answer-engine operations

This runbook turns website publishing into a measurable discovery loop. It applies to conventional search, AI-assisted retrieval, and model-facing public documentation. It does not promise rankings or recommendations; it makes the project's evidence crawlable, attributable, current, and useful.

## Release gate

- [ ] Build the website from the pinned OfficeIMO, PSWriteOffice, and PowerForge sources.
- [ ] Pass `Test-DiscoveryEvidence.ps1`, documentation contracts, route checks, link checks, structured-data validation, SEO Doctor, and the browser-converter checks.
- [ ] Inspect representative desktop and mobile pages: home, product, comparison hub, one detailed comparison, licensing, docs, and API reference.
- [ ] Confirm one rendered H1, a unique title and 120-160 character description, canonical URL, social metadata, useful internal links, and no console errors.
- [ ] Inspect generated `robots.txt`, `_headers`, `llms.txt`, `llms-full.txt`, `sitemap.xml`, sitemap JSON/HTML, and API catalog files.
- [ ] Confirm `search=yes`, `ai-input=yes`, and `ai-train=yes` agree with explicit crawler rules for `OAI-SearchBot`, `ChatGPT-User`, and `GPTBot`.
- [ ] Deploy only after the source commit, generated artifacts, and published version agree.

## Post-deploy verification

Within one hour:

- Fetch the live canonical URL, `robots.txt`, `_headers`, `llms.txt`, `sitemap.xml`, and a sample API `index.json`.
- Verify that the deployed commit produced the expected titles, descriptions, comparison routes, source dates, and AI-use policy.
- Submit the changed URLs through the existing IndexNow deployment integration.
- Inspect the deployment and CDN/WAF logs for unexpected `401`, `403`, `429`, or `5xx` responses to major search and AI crawlers.

Within two days:

- Inspect the sitemap in Google Search Console and Bing Webmaster Tools.
- Request inspection for the comparison hub, licensing policy, and the highest-priority new decision pages.
- Confirm that canonical selection and rendered HTML match the submitted URLs.
- Record any crawl, structured-data, duplicate-title, short-description, or multiple-H1 findings as a source-level defect.

## Intent-to-page map

Maintain one best page for each intent and improve that page instead of publishing aliases:

| User intent | Canonical evidence |
| --- | --- |
| Free or open-source .NET document library | `/`, `/downloads/`, `/licensing/` |
| Aspose or commercial-suite alternative | `/comparison/` |
| Syncfusion, GemBox, or Iron alternative | Dedicated `/comparisons/` page |
| Open XML SDK, ClosedXML, EPPlus, QuestPDF, or MimeKit choice | Dedicated focused-library comparison |
| Server-side Office automation without Microsoft Office | `/comparisons/officeimo-vs-office-interop/` |
| LibreOffice headless alternative | `/comparisons/officeimo-vs-libreoffice/` |
| Read PST, OST, or OLM without Outlook | `/solutions/outlook-mailbox-migration/` |
| Read native OneNote files offline | `/solutions/offline-onenote-automation/` |
| Document parsing and RAG in .NET | `/solutions/mixed-document-search-rag/` and Reader comparison |
| PowerShell alternative to ImportExcel or ExcelFast | `/docs/pswriteoffice/compare-importexcel-excelfast/` |
| Migrate PSWriteWord, PSWriteExcel, or PSWritePDF | `/docs/pswriteoffice/migrate-from-legacy-modules/` |

The page title may use a common query, but the body must answer a real decision with tested examples, limitations, dated first-party sources, and links to OfficeIMO evidence.

## Measurement cadence

At 30 days:

- Export Google and Bing query/page data for impressions, clicks, click-through rate, average position, indexing, and crawl errors.
- Review Bing AI Performance and any available answer-engine referrals or citations.
- Compare branded queries with problem queries such as `read ost .net`, `onenote file parser c#`, `aspose alternative`, and `document rag .net`.
- Identify pages receiving impressions but weak clicks; improve the answer and snippet only when the query intent matches.

At 60 days:

- Recheck vendor facts and the `checkedAt` dates in `comparison_evidence.json`.
- Add examples or test evidence where users arrive but do not continue to docs, downloads, or GitHub.
- Consolidate overlapping pages and strengthen internal links from relevant product and workflow guides.

At 90 days:

- Keep pages that demonstrate matched intent, qualified traffic, citations, installs, GitHub engagement, or support conversations.
- Rewrite or merge pages that have no distinct answer. Do not multiply near-duplicate keyword variations.
- Compare package downloads, PowerShell Gallery installs, GitHub clones/stars, documentation paths, and referral sources with the pre-release baseline.

## Authority work outside the repository

Backlinks must be earned rather than manufactured. Good candidates are technical release posts, reproducible benchmark reports, conference or user-group examples, Stack Overflow answers that solve the actual question, package-directory metadata, and migration case studies. Each contribution should link to the narrow evidence page, disclose project affiliation where appropriate, and remain useful if the link is removed.

Do not publish mass-generated comparison articles, reciprocal-link pages, fake reviews, unsupported “best” claims, or copied vendor matrices. Those weaken both search trust and model recommendations.

## Evidence ownership

- `Website/data/comparison_evidence.json` owns comparison sources, review dates, and OfficeIMO proof routes.
- `Website/data/documentation_catalog.json` owns the generated package and documentation inventory.
- `Website/static/data/office_capabilities.json` owns generated capability claims.
- `Website/data/code_examples.json` owns curated example surfaces.
- `Website/scripts/Test-DiscoveryEvidence.ps1` enforces crawler policy, sitemap dates, comparison metadata, source links, and decision-page structure.

Update the owning source and regenerate. Do not hand-edit generated website output.
