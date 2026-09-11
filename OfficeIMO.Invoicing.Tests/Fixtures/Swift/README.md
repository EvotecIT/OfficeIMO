# IBAN format fixtures

The CSV records the country prefix, BBAN structure, and electronic-format example
from the [SWIFT IBAN Registry, release 102 (June 2026)](https://www.swift.com/swift-resource/9606/download).
It contains the 89 registered prefixes. The `n`, `a`, and `c` fields represent
digits, letters, and alphanumeric characters; `!` denotes a fixed length.

Tests exercise invoice account classification and serialization against these
external format examples. This checks identifier structure and checksum, not
whether an account exists or can receive a payment.
