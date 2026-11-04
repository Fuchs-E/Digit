# Releasing a new version

Digit updates itself from the `release` branch of this repository, which must stay public. Digit reads `version.json` on that branch, compares it with its own version and, when it is newer, downloads the branch's source. Each PC then builds and tests its own `Digit.xlam`.

1. Set `DIGIT_VERSION` in `src/vba/DigitApp.bas` to the new version, for example `1.2.0`.
2. Set the same `version` in `version.json`, and write `notes`. Digit shows the notes in its update prompt, so keep them short and plain.
3. Run the static check, which catches compile errors the self-tests cannot reach (a missing procedure, a misplaced declaration). It must report 0 problems:

   ```
   python tools/check_vba.py
   ```

4. Commit, then push that commit to the `release` branch:

   ```
   git push origin HEAD:release
   ```

Only push to `release` what is ready: every PC that checks for updates is offered it. The installer refuses a download whose `DIGIT_VERSION` differs from the version it was offered, so a mismatch cannot make Digit offer the same update forever.

GitHub's raw file service caches `version.json` for up to 5 minutes, so a new version can take that long to show up.
