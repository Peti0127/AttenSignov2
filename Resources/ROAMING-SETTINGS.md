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

- `Nummer`: `Alles`, `Handy`, `Festnetz`, `Office`, `EDVHotline`
- `MfG`: `MfG0`, `MfG1`, `MfG2`, `MfG3`
- `AutoInsert`: `true` or `false`
- `AutoInsertMode`: `NewMail` or `AllMail`

`MfG0` means no greeting. The other values render the configured greeting and
one blank line before the signature details.

`EDVHotline` is shown in the settings only when the Microsoft 365 profile's
`department` value equals `IT`. It renders `Tel. 05 7999 9999 Mobil {Mobile}`;
if the profile has no mobile number, only `Tel. 05 7999 9999` is rendered. If a
user's department later changes away from IT, the option is no longer offered
in Settings. The automatic runtime treats an already saved `EDVHotline` value
as authoritative so that an older cached profile can't incorrectly replace it
with the standard `Alles` phone line.

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

The event runtime intentionally avoids `async`/`await` and the conditional
operator so it can load in classic Outlook builds that use the older event
runtime. After changing the manifest, remove the installed add-in, sideload the
new manifest, and restart Outlook so the launch-event registration and runtime
bundle are refreshed.

No SQL database, API endpoint, database credentials, or additional Microsoft
Graph permission is required for these preferences.
