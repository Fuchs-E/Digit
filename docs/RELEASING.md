# Releasing a new version

Digit updates itself from GitHub releases of this repository, which must stay public. A release holds only source code: each PC builds and tests its own `Digit.xlam` when it updates.

1. Set `DIGIT_VERSION` in `src/vba/DigitApp.bas` to the new version, for example `1.2.0`.
2. Commit and push.
3. Create a release whose tag is the same version with a `v` in front (`v1.2.0`), on the commit from step 2. Write a few lines in the release notes, because Digit shows them in its update prompt. Do not mark it as a pre-release or a draft.

The version in the code and the tag must match. If they do not, the installer refuses the release, so Digit cannot end up offering the same update forever.
