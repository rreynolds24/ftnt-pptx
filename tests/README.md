# Offline Regression Tests

Run from the package root:

```sh
python -S -m unittest discover -s tests -v
```

`-S` disables site-package loading. The 56-test suite uses only the Python standard library and sanitized synthetic OOXML and SVG fixtures. It tests non-destructive planning/application, text/notes/hyperlink/chart-value preservation, theme/run formatting, namespace survival, exact palette rules, explicit table/shape operations, per-instance image replacement, logo minimum/contrast/confirmation boundaries, input/asset/catalog hashes, idempotence, protected opaque parts, archive/XML safety, local icon import staging, and CLI execution.

Tests are not proof of Microsoft Office renderer parity or of any Fortinet product capability. Separate integration checks apply the tool to the actual completed source slide, render the result locally with LibreOffice, and inspect Inter/font and overflow evidence. The reusable builder is also rendered and visually inspected.

All output files are created in temporary directories. No network or customer environment access is required.
