# Automatic insertion deployment

Automatic insertion is an Outlook launch event, so replacing only the web
files isn't enough. Complete these steps in order.

1. Upload every file from the package to the GitHub Pages `Resources` folder,
   including `autorun.html` and `autorun.js`.
2. Wait until GitHub Pages has published the new files.
3. Remove the currently sideloaded Attensam add-in from Outlook.
4. Close Outlook completely.
5. Sideload the included `manifest.xml` version `1.3.2.0`.
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

Manifest version `1.3.2.0` includes both the mobile settings command and the mobile
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
