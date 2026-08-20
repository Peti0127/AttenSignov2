# Automatic insertion deployment

Automatic insertion is an Outlook launch event, so replacing only the web
files isn't enough. Complete these steps in order.

1. Upload the hosted add-in files from the package to the GitHub Pages
   `Resources` folder, including `autorun.html`, `autorun.js`, and
   `autorun-classic.js`. The Markdown documentation and source-only
   `msal-shim.js` don't need to be hosted. The well-known JSON file has a
   separate root location described below.
2. Wait until GitHub Pages has published the new files.
3. Remove the currently sideloaded Attensam add-in from Outlook.
4. Close Outlook completely.
5. Sideload the included `manifest.xml` version `1.4.1.0`.
6. Restart Outlook.
7. Open a compose window and open the task pane once. Wait for the status
   `Microsoft-365-Profil wurde automatisch geladen.` This stores the profile
   and template needed by the lightweight event runtime.
8. Enable `Signatur automatisch einfügen`, select a mode, and close that
   compose window.
9. Create another message to test the launch event.

`Einfügen bei neuen Emails` applies only to compose type `newMail`.
`Einfügen bei allen Emails` also applies to replies, reply-all messages, and
forwards. Opening an existing draft doesn't trigger `OnNewMessageCompose`.

If the task pane reports that automatic signature data couldn't be saved, the
event runtime intentionally skips insertion instead of inserting stale data.

## Outlook Mobile

Manifest version `1.4.1.0` includes both the mobile settings command and the mobile
`OnNewMessageCompose` launch event.

1. Use Outlook for Android or iOS version `4.2352.0` or later.
2. For a company rollout, deploy `manifest.xml` as an integrated app in the
   Microsoft 365 admin center. After deployment, allow time for Outlook Mobile
   to receive the add-in.
3. In Outlook Mobile, open a received email, open the Apps/add-ins menu, and
   select `Einstellungen` under the Attensam add-in.
4. Wait for `Profil und Einstellungen sind bereit.` This loads the Microsoft
   365 profile and stores the template data required by the mobile event
   runtime.
5. Enable automatic insertion, choose the desired insertion mode, and tap
   `Fertig`.
6. Create a new message to test the automatic signature.

Outlook Mobile doesn't expose normal add-in task panes while composing.
Therefore the settings command is intentionally available while reading a
message, and signature insertion happens automatically in compose mode. On
iOS, composing through the Share action doesn't trigger the launch event.

## Send As / Im Auftrag von

Version `1.4.1.0` reads the current compose item’s From field. If its address
belongs to a different user, the add-in inserts `Im Auftrag von` followed by
that user’s name, `extensionAttribute10`, and `extensionAttribute11`. Changing
the From field updates an automatically inserted signature through the
`OnMessageFromChanged` event.

This lookup requires additional Microsoft Graph configuration:

1. In the Microsoft Entra app registration used by `clientId`, add the
   delegated Microsoft Graph permission `User.Read.All`.
2. Select `Grant admin consent` for the tenant. Individual users can't consent
   to this permission themselves.
3. Upload `autorun-classic.js` to the GitHub Pages `Resources` directory with
   the other web files.
4. For classic Outlook on Windows, publish the included
   `microsoft-officeaddins-allowed.json` at the origin-root URL
   `https://peti0127.github.io/.well-known/microsoft-officeaddins-allowed.json`.
   This is the domain root, not the `AttenSignov2/Resources` directory. A
   GitHub Pages user-site repository or a custom domain may be needed to serve
   that root path.

If the directory lookup is temporarily unavailable, the add-in still uses the
display name supplied by Outlook, but the two extension attributes can't be
included until `User.Read.All` and the event-runtime authorization are working.

## Live settings updates

Signatures inserted by version `1.4.1.0` contain a non-visible Attensam marker.
When a user changes a preference in the desktop compose settings page, the
add-in checks the current message for that marker. If present, it rebuilds and
replaces the signature immediately, including the current From identity. If no
Attensam signature is present, changing a setting doesn't insert one.

Signatures inserted by an older add-in version don't contain the marker and
therefore can't be detected for live updating. Insert the signature once with
version `1.4.1.0` to enable this behavior. Outlook Mobile opens settings in
read mode, so mobile changes apply to the next compose item.
