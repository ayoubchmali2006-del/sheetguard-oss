# Contributing

Thank you for your interest in SheetGuard OSS.

## Development setup
Python 3.10 or newer is required. No runtime dependencies are required.

```bash
python -m pip install -e .
python -m unittest discover -s tests -v
```

## Proposing changes
1. Open an issue describing the problem or enhancement before a large change.
2. Create a branch and add tests that reproduce any bug.
3. Keep the tool read-only and avoid introducing unnecessary dependencies.
4. Run the test suite locally.
5. Submit a pull request describing behavior, limitations and tests.

Never upload private Excel workbooks, credentials, customer records or proprietary materials into public issues or pull requests. Create synthetic reproductions instead.

## Scope
The tool reports structural OOXML signals only. It does not execute Excel macros, edit workbooks, recalculate formulas or certify security.

## Code of conduct
Be respectful, constructive and mindful of privacy.
