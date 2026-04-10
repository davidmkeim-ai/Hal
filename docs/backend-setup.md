# HAL Backend Setup

HAL now includes a small Node/Express backend scaffold for local development, optional Microsoft sign-in, and future Microsoft Graph meeting creation.

## Files

- [package.json](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\package.json)
- [.env.example](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\.env.example)
- [docs/entra-app-registration.md](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\docs\entra-app-registration.md)
- [server/index.js](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\server\index.js)
- [server/config.js](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\server\config.js)
- [server/msal.js](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\server\msal.js)
- [server/graph.js](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\server\graph.js)

## Intended local flow

1. Install Node.js 20 or newer.
2. Run `npm install`.
3. Run `npm run dev`.
4. Open `http://localhost:4000/index.html`.
5. Add `.env` later when you are ready to test Microsoft sign-in.

## What the backend currently does

- Serves the HAL frontend
- Supports a local-first development loop with no Microsoft setup required
- Starts Microsoft sign-in with MSAL
- Stores the signed-in user session
- Creates calendar-backed Teams meetings through Microsoft Graph

## Important current limitation

This scaffold currently uses a standard Microsoft Entra web sign-in redirect flow.

That is enough to start local development and Graph testing, but it is not yet the full Teams SSO pattern described by Microsoft for tab apps. A later upgrade can replace the redirect-first approach with Teams `authentication.getAuthToken()` plus server-side token exchange.

## Next likely backend step

Add a real database so HAL notes and checklists sync across devices and across Teams sessions.
