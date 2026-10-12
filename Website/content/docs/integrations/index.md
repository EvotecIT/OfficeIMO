---
title: "Solve reporting tasks with PowerShell"
description: "Turn Windows events and HTML supplier quotes into useful documents, then prepare report emails with readable PowerShell examples on each page."
layout: docs
---

Getting data into PowerShell is often the easy part. The next person needs something they can read, filter, edit or approve. These articles walk through that last step, using PSWriteOffice with the modules that collect or share the data.

## Hand over an investigation without losing the event details

A screenshot of Event Viewer gives the next investigator too little context. A long list gives them too much to read at once. [Turn Windows events into an investigation handoff](/docs/integrations/event-brief/) shows how to collect a bounded snapshot with PSEventViewer, write a short Word/PDF brief, and keep the full messages in Excel.

## Compare quotes without copying the same table twice

An HTML price list needs to become a purchasing decision. [Compare supplier quotes](/docs/integrations/supplier-review/) reads the offers with PSParseHTML, keeps prices numeric, and builds an editable review alongside the workbook. The example explains why the cheapest offer and the quickest delivery are different choices.

## Check the email as carefully as the report

The right report is no help if the email has the wrong attachment. [Review a report email before sending it](/docs/integrations/report-email/) uses Mailozaurr to compose a message with a PDF and workbook, save it locally for inspection, and add SMTP delivery when you're ready.

Each article shows the code in the page, explains the data it works with and links the resulting documents. Run the blocks in order or adapt the parts you need to an existing task. The sample downloads use illustrative data; the event collection code reads your Windows log, and the email send step is separate from local message creation.
