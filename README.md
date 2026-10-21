# Hindi text for ElevenLabs

Drop English/Hindi Word (.docx) or XLIFF (.xliff / .xlf) files on the page and get the cleaned Hindi text,
split into 2-3 sentence blocks with pause tags, ready to copy into ElevenLabs or Cartesia.

Each output starts with the passage reference (e.g. `निर्गमन अध्याय 30, वचन 34 से 38`).
If a file has none, it is copied from another file of the same passage dropped
together (same name apart from the `-s2`, `-s3` section part).

Everything runs in the browser. Files are never uploaded. Results are remembered on
your own computer until you press **Clear all**.

**Changing the rules:** edit `prompt.txt` (word changes like `साँप to सांप`, the scripture
example, and `pause 5 seconds after: ...` lines). The site updates a minute after you save.
