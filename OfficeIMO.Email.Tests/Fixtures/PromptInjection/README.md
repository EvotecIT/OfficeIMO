# Email instruction fixtures

These ten synthetic EML fixtures come from [cyb3rmik3/prompt-injection-email-samples](https://github.com/cyb3rmik3/prompt-injection-email-samples), commit `88293a1543f531174ba915717bd9e1f6090a5563`. The original MIT license is included as `LICENSE`.

They cover visible, concealed HTML and inline Base64 requests for prompt disclosure, private-data transfer and tool discovery, plus a benign control. Tests open them as untrusted data and never follow URLs or execute their instructions. Expected classifications are test inputs; fixture headers are not detector signals.
