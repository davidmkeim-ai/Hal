# HAL Microsoft Teams Setup

HAL is currently scaffolded as a personal tab app for Microsoft Teams.

## What HAL needs

1. A public HTTPS URL where [index.html](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\index.html) is hosted.
2. A Microsoft Entra app registration for sign-in and Microsoft Graph access.
3. A Teams app ID for the app package in [appPackage/manifest.json](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\appPackage\manifest.json).

For the Entra app registration details, use [docs/entra-app-registration.md](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\docs\entra-app-registration.md).

## Recommended delegated Microsoft Graph permissions

- `User.Read`
- `Calendars.ReadWrite`
- `offline_access`
- `openid`
- `profile`

## Recommended meeting creation approach

For HAL, the better default is to create a calendar event with Microsoft Teams enabled instead of creating a standalone `onlineMeeting`.

Why:

- A calendar-backed event shows up in Outlook and Teams calendars.
- It fits your "upcoming call" workflow better.
- Microsoft Graph's standalone `onlineMeeting` API creates meetings that do not appear on calendars.

## Suggested build sequence

1. Move HAL from static local storage to a small backend with a database.
2. Add Microsoft sign-in.
3. Exchange a signed-in user token for Microsoft Graph access.
4. Implement calendar-backed Teams meeting creation from call-prep entries.
5. Attach exported drafts to Word, Excel, or PowerPoint workflows.

## Notes

- Teams tabs must be hosted on HTTPS.
- Microphone access depends on the browser or Teams host granting media permissions.
- The current app package is a scaffold, not a production-ready package.
