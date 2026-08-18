# Outlook roaming settings

The add-in stores each user's signature choices in the user's Microsoft 365
mailbox with `Office.context.roamingSettings`.

## Stored record

Key: `attensam.signature.settings.v2`

```json
{
  "Nummer": "Alles",
  "MfG": "MfG1",
  "updatedAt": "2026-08-18T12:00:00.000Z"
}
```

Allowed values:

- `Nummer`: `Alles`, `Handy`, `Festnetz`, `Office`
- `MfG`: `MfG0`, `MfG1`, `MfG2`, `MfG3`

`MfG0` means no greeting. The other values render the configured greeting and
one blank line before the signature details.

## Local cache

`localStorage` remains only as a per-user cache. Outlook initializes roaming
settings when the task pane opens and may expose the previous in-memory value
after an immediate page navigation. The `updatedAt` timestamp lets the add-in
use the newest mailbox or cache record. The mailbox copy remains the persistent,
cross-client source of truth.

No SQL database, API endpoint, database credentials, or additional Microsoft
Graph permission is required for these preferences.
