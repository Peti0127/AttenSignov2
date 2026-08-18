# SQL integration contract

The current `preferences.js` uses per-user `localStorage` as a temporary adapter. Replace only `getSettings()` and `saveSettings()` when the backend is available. Both methods already use the final SQL/API record shape.

Recommended table:

```sql
CREATE TABLE UserSignatureSettings (
  EntraUserObjectId VARCHAR(64) NOT NULL PRIMARY KEY,
  Nummer NVARCHAR(16) NOT NULL CONSTRAINT DF_UserSignatureSettings_Nummer DEFAULT N'Alles',
  MfG VARCHAR(8) NOT NULL CONSTRAINT DF_UserSignatureSettings_MfG DEFAULT 'MfG1',
  UpdatedAtUtc DATETIME2 NOT NULL DEFAULT SYSUTCDATETIME(),
  CONSTRAINT CK_UserSignatureSettings_Nummer
    CHECK (Nummer IN (N'Alles', N'Handy', N'Festnetz', N'Office')),
  CONSTRAINT CK_UserSignatureSettings_MfG
    CHECK (MfG IN ('MfG1', 'MfG2', 'MfG3'))
);
```

Stored values and rendered output:

| Column | Stored value | Signature output |
| --- | --- | --- |
| `Nummer` | `Alles` | `Tel.: {Phone} Mobil: {Mobile}` with automatic fallback if one number is empty |
| `Nummer` | `Handy` | `Mobil {Mobile}` |
| `Nummer` | `Festnetz` | `Tel.: {Phone}` |
| `Nummer` | `Office` | The fixed company office number from `CONFIG.officeNumber` |
| `MfG` | `MfG1` | `Mit freundlichen Grüßen` |
| `MfG` | `MfG2` | `Freundliche Grüße` |
| `MfG` | `MfG3` | `LG` |

`template.html` does not contain an `{MfG}` placeholder. `taskpane.js` creates
the complete greeting paragraph only when `MfG` equals `MfG1`, `MfG2`, or
`MfG3`. Any other value inserts no greeting.

Recommended authenticated endpoints:

- `GET /api/user-settings` → `{ "Nummer": "Alles", "MfG": "MfG1" }`
- `PUT /api/user-settings` with `{ "Nummer": "Handy", "MfG": "MfG2" }`

The API must derive `EntraUserObjectId` from the validated Microsoft Entra access token. It must not accept a user ID supplied by the browser. Database credentials must remain server-side.
