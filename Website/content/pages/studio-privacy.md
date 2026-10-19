---
title: "OfficeIMO Studio privacy"
description: "Learn how OfficeIMO Studio processes documents on your device and stores local application data on Apple, Windows, and Linux devices."
meta.seo_title: "OfficeIMO Studio privacy and local document data"
layout: page
slug: studio/privacy
---

OfficeIMO Studio opens and processes documents on your device. Opening, reading,
editing, converting, and saving a document does not send it to Evotec. Studio has
no analytics, telemetry, or automatic update checks. Store editions are updated by
the Microsoft Store or Apple's App Store.

For the OfficeIMO website, browser tools, and libraries, see the
[OfficeIMO privacy policy](/privacy/).

## Native app for iPhone, iPad, and Mac

The native Apple app does not collect data for Evotec, track you, show ads, or
require an account. It does not include the document assistant described below.
Document processing uses OfficeIMO engines in the app. Scanning and text
recognition use Apple's on-device frameworks; the native app does not download
Tesseract language files or external OCR programs.

You choose documents through the system document browser. Documents remain in
the location you select, which may be on your device or with a storage provider
such as iCloud Drive. That provider's settings and privacy policy govern its
storage and synchronization. Camera access is used when you scan a document.
The photo picker gives Studio access to the photos you select.

Studio stores reading preferences and recovery data in its private application
container. Recovery copies contain document contents and support restoring
unsaved work. Recovery snapshots expire after 30 days. Working copies, scans,
conversion results, and files prepared for sharing can also remain in the app's
temporary storage until cleanup. Deleting an original document does not
necessarily delete its recovery or temporary copies immediately.

Sharing, printing, and opening a link use the system services you choose. Those
services may transfer the selected data to another app or provider. Apple also
controls App Store transactions, operating-system diagnostics, device backups,
and synchronization under your Apple settings and its privacy policy.

## Windows and Linux application data

Studio stores its data in a folder for your user account:

- **Windows and Linux:** `%LOCALAPPDATA%\OfficeIMO\Studio` (on Linux, the
  equivalent local application data folder)

Earlier Avalonia-based macOS source builds use
`~/Library/Application Support/OfficeIMO/Studio`. This section describes that
desktop implementation as well as Windows and Linux.

That folder contains:

- **Settings:** your preferences, including the document assistant provider, model,
  and endpoint you chose. API keys are not saved in settings.
- **Recent documents and sessions:** file paths or storage-provider bookmarks for up
  to 32 recent documents. Entries expire after 30 days.
- **Reading positions:** where you left off in up to 64 documents, keyed by a hash of
  the file path.
- **Recovery copies:** when recovery is enabled, copies of documents you are editing,
  so unsaved work can be restored. **These copies contain document contents.**
  Session snapshots expire after 30 days. Workflow recovery copies are limited to
  1 GiB in total.
- **Saved signatures:** signature images you create.
- **Document assistant sign-in:** sign-in tokens for ChatGPT or GitHub Copilot, if you
  connect them.
- **Diagnostics:** local log files of up to 2 MB each, keeping the 5 most recent.
  They contain event codes, operating system and runtime details, exception types,
  and stack traces with file paths removed. They exclude document contents,
  document names and paths, and exception messages. Diagnostics are never sent
  anywhere automatically.

Studio also writes short-lived temporary files, such as conversion previews, to
your system's temporary folder.

Settings provides controls for document history and recovery data. **Uninstalling
Studio does not remove this folder.** Delete it yourself to remove all Studio data
from your device.

## Optional document assistant in the desktop implementation

The document assistant uses the provider and model you select. Remote processing
requires enabling the remote-processing control. When you ask a question using
a remote provider, Studio sends your question and selected document evidence to
that provider. Follow-up requests can include up to three earlier exchanges from
the conversation. Provider privacy policies, account terms, retention policies,
and usage limits apply.

Available connections include ChatGPT, an OpenAI-compatible endpoint, GitHub
Copilot, and a local-model endpoint. Remote endpoints must use HTTPS. Local-model
connections require a loopback HTTP(S) endpoint. Check the local server's behavior
before sending sensitive content; a local gateway can forward requests to another
service.

Sign-in and model discovery communicate with the selected service without
sending document evidence. The service can return account details, such as your
email address and plan, which Studio shows in the connection settings. Studio
retains supported account credentials in its application-specific authentication
files or session memory. API credentials entered for compatible endpoints remain in
session memory; non-secret provider, model, and endpoint choices can be retained in
settings. Protect the device and its application-data backups. Use Sign out to end
the selected Studio connection.

## Desktop OCR language data

Searchable-PDF and text-recognition workflows use Tesseract language data. When
the download option is selected (it is on by default), Studio downloads the
language files you need from the Tesseract project on GitHub
(`raw.githubusercontent.com`) and checks them against known SHA-256 hashes. This
request reveals your IP address to GitHub; no document content is sent. The files
are kept in `OfficeIMO\Ocr` in the local application data folder for later use.
Clear the download option to use only language data you install yourself.

## External tools and support

The direct desktop edition can use an optional Tesseract installation for OCR
and operating-system command-line tools for printer queue delivery. These tools
process the inputs supplied to the operation and have their own licenses and
privacy behavior. The native Apple app does not discover or download external
OCR programs or start printer command-line tools. Document conversions use
OfficeIMO engines in the application.

Links in Studio, such as sign-in pages, help pages, and links inside documents,
open in your default web browser.

Files, screenshots, or logs attached to a public GitHub issue are public. Remove
personal information and credentials before sharing an example. For private
support or privacy questions, contact [contact@evotec.pl](mailto:contact@evotec.pl).

OfficeIMO's MIT license and public-source policy cover the software and public
project material. They do not make your documents, credentials, or local
application data public.
