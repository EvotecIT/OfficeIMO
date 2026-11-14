# OpenPGP fixtures

`pgpy-messages.json` holds PGP/MIME and inline OpenPGP messages produced by PGPy 0.6.0, an
OpenPGP implementation independent of BouncyCastle. Alice has an RSA 3072 primary key with an RSA
encryption subkey, protected with a passphrase; Bob has an Ed25519 primary key with a Curve25519
encryption subkey. They are throwaway test keys. Key blocks and messages are base64 so their exact
bytes, including CRLF line endings, survive.
