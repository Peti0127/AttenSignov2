# Outlook roaming settings

The add-in stores each user's signature choices in the user's Microsoft 365
mailbox with `Office.context.roamingSettings`.

## Stored record

Key: `attensam.signature.settings.v2`

```json
{
  "Nummer": "Alles",
  "MfG": "MfG1",
  "CustomGreeting": "",
  "AutoInsert": false,
  "AutoInsertMode": "NewMail",
  "InsertTitleBefore": false,
  "InsertTitleAfter": false,
  "MobileUsage": false,
  "Confidentiality": false,
  "updatedAt": "2026-08-18T12:00:00.000Z"
}
```

Allowed values:

- `Nummer`: `Alles`, `Handy`, `Festnetz`, `Office`, `EDVHotline`
- `MfG`: `MfG0`, `MfG1`, `MfG2`, `MfG3`, `MfGCustom`
- `CustomGreeting`: the user's custom greeting text, up to 200 characters
- `AutoInsert`: `true` or `false`
- `AutoInsertMode`: `NewMail` or `AllMail`
- `InsertTitleBefore`: `true` or `false`
- `InsertTitleAfter`: `true` or `false`
- `MobileUsage`: `true` or `false`
- `Confidentiality`: `true` or `false`

`MfG0` means no greeting. The other values render the configured greeting and
one blank line before the signature details.

`MfGCustom` renders the normalized plain text stored in `CustomGreeting`. The
desktop and mobile settings pages show the custom text field only while this
mode is selected. HTML entered by a user is escaped before rendering.

`InsertTitleBefore` inserts `extensionAttribute10` before the first name, and
`InsertTitleAfter` inserts `extensionAttribute11` after the last name. Each
checkbox is shown only when the corresponding Microsoft 365 profile attribute
contains a value. The add-in normalizes the surrounding spaces, so titles are
separated cleanly from the name regardless of whitespace stored in Graph.

`MobileUsage` appends “Diese E-Mail wurde über Outlook Mobile versendet.” in
12 pt beneath the signature. `Confidentiality` appends “Diese E-Mail ist
vertraulich.” in 9 pt. When both are enabled, the mobile notice is rendered
first and the confidentiality notice second.

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
The same render-data record also contains a timestamped copy of all preferences.
Every settings save updates both roaming records in one `saveAsync` operation.
The automatic runtime selects the newest timestamped copy, preventing an older
`Nummer` value from falling back to `Alles` while the task pane already shows
the new selection.

The render-data record also contains the public Entra client and tenant IDs
needed by the event runtime. When the Outlook From address differs from the
signed-in profile, the runtime requests a `User.Read.All` token silently and
loads the selected sender's complete signature profile, including company,
address, job title, telephone numbers, and both extension attributes. The
signature uses these From-user details while rendering the name as
`signed-in user (i.A. From user)`. The token itself is never written to roaming
settings.

The add-in compares the normalized domain after the final `@` in both email
addresses before requesting the From user's profile. Exact same-domain matches
are allowed; different or missing domains disable delegated rendering and use
the signed-in user's normal signature.

Rendered signatures include the `attensam-signature-root` element and a hidden
`ATTENSAM-SIGNATURE-V2` text marker. After a preference is saved, the desktop
settings page uses Outlook's native signature replacement when available. On
older clients, it replaces the marked element through the compose body API.

- `NewMail` inserts only when `getComposeTypeAsync()` returns `newMail`.
- `AllMail` inserts for new messages, replies, reply-all messages, and forwards.
- Editing an existing draft doesn't trigger `OnNewMessageCompose`.

The automatic feature uses APIs introduced after Mailbox requirement set 1.5.
Outlook Mobile exposes the required launch-event, `setSignatureAsync`, and
`getComposeTypeAsync` APIs as documented mobile exceptions. The
`VersionOverridesV1_1` configuration in `manifest.xml` registers the event for
both desktop and mobile. Upload
`autorun.html`, `autorun.js`, and `autorun-classic.js` to the same GitHub Pages `Resources` directory
as the task-pane files before installing the updated manifest.

On Outlook Mobile, open the `Einstellungen` command while reading a message.
The full-screen `mobile-settings.html` page retrieves the same Microsoft Graph
profile and template as the desktop task pane, then refreshes
`attensam.signature.render-data.v1`. This means a user can initialize and
manage automatic insertion entirely from Outlook Mobile.

The event runtime intentionally avoids `async`/`await` and the conditional
operator so it can load in classic Outlook builds that use the older event
runtime. After changing the manifest, remove the installed add-in, sideload the
new manifest, and restart Outlook so the launch-event registration and runtime
bundle are refreshed.

No SQL database, API endpoint, or database credentials are required. Reading
another sender's full directory profile requires the delegated Microsoft Graph
permission `User.Read.All` with tenant administrator consent.
