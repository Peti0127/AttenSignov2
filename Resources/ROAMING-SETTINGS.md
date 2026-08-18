# Outlook roaming settings

The add-in stores each user's signature choices in the user's Microsoft 365
mailbox with `Office.context.roamingSettings`.

## Stored record

Key: `attensam.signature.settings.v2`

```json
{
  "Nummer": "Alles",
  "MfG": "MfG1",
  "AutoInsert": false,
  "AutoInsertMode": "NewMail",
  "updatedAt": "2026-08-18T12:00:00.000Z"
}
```

Allowed values:

- `Nummer`: `Alles`, `Handy`, `Festnetz`, `Office`
- `MfG`: `MfG0`, `MfG1`, `MfG2`, `MfG3`
- `AutoInsert`: `true` or `false`
- `AutoInsertMode`: `NewMail` or `AllMail`

`MfG0` means no greeting. The other values render the configured greeting and
one blank line before the signature details.

## Local cache

`localStorage` remains only as a per-user cache. Outlook initializes roaming
settings when the task pane opens and may expose the previous in-memory value
after an immediate page navigation. The `updatedAt` timestamp lets the add-in
use the newest mailbox or cache record. The mailbox copy remains the persistent,
cross-client source of truth.

## Automatic insertion

The task pane stores the profile and unrendered signature template under
`attensam.signature.render-data.v1`. The lightweight event runtime renders that
data with the user's current settings whenever Outlook creates a compose item.

- `NewMail` inserts only when `getComposeTypeAsync()` returns `newMail`.
- `AllMail` inserts for new messages, replies, reply-all messages, and forwards.
- Editing an existing draft doesn't trigger `OnNewMessageCompose`.

The automatic feature requires Mailbox requirement set 1.10 and the
`VersionOverridesV1_1` launch-event configuration in `manifest.xml`. Upload
`autorun.html` and `autorun.js` to the same GitHub Pages `Resources` directory
as the task-pane files before installing the updated manifest.

No SQL database, API endpoint, database credentials, or additional Microsoft
Graph permission is required for these preferences.
