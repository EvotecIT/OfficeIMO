# Email authentication fixtures

- `rfc8463-appendix-a.eml`: the signed message from RFC 8463 Appendix A.3, with the
  RSA-SHA256 and Ed25519-SHA256 signatures and the key records from Appendix A.2.
- `arc-test-suite-validation.json`: the validation cases from the ARC test suite
  (https://github.com/ValiMail/arc_test_suite, commit f137dcb9d6d5baeef1310024ce9ccca94a9a92c8,
  `tests/arc-draft-validation-tests.yml`), converted from YAML to JSON without changing messages,
  key records or expected results. Licensed under the MIT licence:

  > Copyright <2017> ValiMail Inc.
  >
  > Permission is hereby granted, free of charge, to any person obtaining a copy of this software and
  > associated documentation files (the "Software"), to deal in the Software without restriction,
  > including without limitation the rights to use, copy, modify, merge, publish, distribute,
  > sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is
  > furnished to do so, subject to the following conditions:
  >
  > The above copyright notice and this permission notice shall be included in all copies or
  > substantial portions of the Software.
  >
  > THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT
  > NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
  > NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM,
  > DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT
  > OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

- `dkimpy-signed.json`: messages signed by dkimpy 1.1.8 with throwaway RSA 2048, RSA 1024 and
  Ed25519 keys (only the public key records are kept), some modified after signing. Messages are
  base64 so their exact bytes survive. They check canonicalization against an independent signer.
