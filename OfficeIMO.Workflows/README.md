# OfficeIMO.Workflows

`OfficeIMO.Workflows` is the reusable local orchestration layer for OfficeIMO document jobs. It composes the existing first-party conversion and PDF APIs behind typed requests, bounded execution, cooperative cancellation, collision policies, atomic output publication, and post-write validation.

The package does not add a second document or PDF engine. Applications such as OfficeIMO Studio, command-line tools, and services can share this workflow contract while keeping their user-interface and hosting code thin.
