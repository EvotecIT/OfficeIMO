---
title: "Privacy"
description: "How the OfficeIMO website, browser tools, Studio, and libraries handle your documents and data. Documents stay on your device unless you choose to send them somewhere."
layout: page
slug: privacy
meta.seo_title: "OfficeIMO privacy policy"
meta.social_card_badge: "Privacy"
---

_Last updated: 6 October 2026_

OfficeIMO is built so that your documents stay with you. The browser tools run in your browser, Studio runs on your computer, and the libraries run inside your own code. None of them upload your documents to Evotec.

This page explains what data each part of OfficeIMO handles, who receives it, and what you can do about it.

## Who is responsible

OfficeIMO is developed by **Evotec Services sp. z o.o.** ("Evotec", "we"), a company registered in Poland. We are the controller of the limited personal data described here. Registration details are on [Evotec's company page](https://evotec.xyz/company-facts/).

For privacy questions or requests, email [support@evotec.pl](mailto:support@evotec.pl).

## At a glance

| Part of OfficeIMO | Your documents | Data that leaves your device |
|---|---|---|
| Website (officeimo.com) | Not involved | Standard web request data, handled by our hosting providers |
| Browser tools | Processed in your browser, never uploaded | None beyond loading the page |
| OfficeIMO Studio | Processed on your computer | Only what you choose to send to an assistant provider, plus optional OCR language downloads |
| .NET libraries, CLI, and PowerShell module | Processed by your code | Nothing, unless your code calls a network feature |
| Downloads and app stores | Not involved | Download requests, handled by the store or download host |

## The website

**No tracking.** officeimo.com does not use analytics, advertising, tracking pixels, or third-party cookies, and it sets no cookies of its own. Fonts, scripts, and search run from the site itself, so pages don't load third-party resources.

**Saved preferences.** Your theme choice (light or dark) and background style are saved in your browser's local storage, so the site looks the same on your next visit. They never leave your browser. You can remove them by clearing site data in your browser.

**Hosting.** The website is hosted on GitHub Pages and delivered through Cloudflare. Like any web server, these providers receive the standard information your browser sends with each request: your IP address, the page requested, the time, and your browser's user agent. They process it to deliver the site and protect it from abuse, under their own privacy policies:

- [GitHub Privacy Statement](https://docs.github.com/site-policy/privacy-policies/github-general-privacy-statement)
- [Cloudflare Privacy Policy](https://www.cloudflare.com/privacypolicy/)

**Links to other sites.** Pages link to GitHub, NuGet, the PowerShell Gallery, Discord, and app stores. Those sites have their own privacy policies.

## Browser tools

The [browser tools](/convert/) convert, organize, and inspect documents entirely in your browser using WebAssembly. Files you open are read from your device by the page and processed in memory. **They are not uploaded to Evotec or anyone else.** Results are handed back to you as a download created in your browser.

The only other requests the tools make are for the sample documents bundled with the site, when you choose to try one.

## OfficeIMO Studio

Studio processes documents on your computer. Opening, editing, converting, and saving a document does not send it to Evotec, and Studio has no analytics, telemetry, or automatic update checks.

Studio sends data over the network only for features you use:

- **Document assistant (optional):** when you ask a question using a hosted provider (ChatGPT, an OpenAI-compatible service, or GitHub Copilot), Studio sends your question and the selected document text to that provider. You must first allow this in Studio. The provider's own terms and privacy policy apply.
- **OCR language data (optional):** when you run text recognition, Studio can download language files from the Tesseract project on GitHub. This download is on by default and reveals your IP address to GitHub. The Mac App Store edition does not download OCR files.

Studio also keeps settings, recent files, recovery copies, and local diagnostics on your computer. Recovery copies can contain document contents.

The [Studio privacy details](/studio/privacy/) list exactly what Studio stores, where it stores it, and how to remove it.

## .NET libraries, CLI, and PowerShell module

The OfficeIMO NuGet packages, the `OfficeIMO.Tool` command-line tool, and the PSWriteOffice PowerShell module contain no telemetry and do not check for updates. They run inside your application or script and process documents there.

Some features make network requests when your code uses them, for example:

- loading images or stylesheets from web addresses in HTML
- connecting to Google Workspace, Confluence, or the Polish KSeF invoicing service
- downloading OCR language data

These requests go directly from your environment to the service you chose. Evotec does not receive them.

## Downloads and app stores

- **Microsoft Store, Mac App Store, and App Store:** when you install OfficeIMO Studio from a store, the store operator handles the download and your account under its own privacy policy. Stores may share aggregated install statistics with us. They don't share your identity.
- **Direct downloads:** Windows installers are served from downloads.officeimo.com, operated by Evotec. It records download requests (time, file, IP address, and browser details) to deliver files, protect the service from abuse, and count downloads. We keep these records for a limited period. Other files are downloaded from GitHub Releases.
- **Package managers:** NuGet, the PowerShell Gallery, and WinGet handle package downloads under their own policies.

## Support and GitHub

If you open an issue or discussion on GitHub, anything you post is public. That includes files, screenshots, and logs. Remove personal information, credentials, and confidential document content before sharing an example.

If you email us, we use your message and address only to answer you and keep a record of the conversation.

## Why we process data, and your rights

We process the limited personal data above (web request data, download records, and support messages) based on our legitimate interests: running and securing the website and downloads, and answering your questions.

Under the EU General Data Protection Regulation (GDPR) you can ask us to:

- access, correct, or delete your personal data
- restrict how we use it
- stop processing it (object)
- provide a copy of it

Email [support@evotec.pl](mailto:support@evotec.pl) with your request. You can also complain to the Polish data protection authority, the [President of the Personal Data Protection Office (UODO)](https://uodo.gov.pl/en), or to the authority in your own country.

## Changes to this policy

We update this page when OfficeIMO's data handling changes. The date at the top shows the latest revision. Significant changes are noted in the release notes.
