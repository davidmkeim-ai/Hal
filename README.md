# HAL

A lightweight scaffold for a personal assistant app that helps capture notes, organize ideas, track daily tasks, and prepare structured drafts for later export into documents and Microsoft workflows.

## What this starter includes

- Voice or typed capture using the browser's Web Speech API when available
- Entry categories for ideas, checklists, call prep, brain dumps, and follow-ups
- Local repository with search and category filters
- Daily checklist with persistent local storage
- Local-first follow-up queue for ideas that should become actions later
- Draft-style controls for shaping notes into usable formats
- Selected-note workspace with HAL prompts and note signals
- Draft generator that turns rough notes into a more structured document outline
- Plain text and Markdown exports for reusable drafts
- Optional OpenAI-backed interpretation layer for smarter classification and formatting
- Local memory for prompt history, correction tracking, and learned title preferences
- Local backup and restore tools, plus timestamped snapshot files in the `backups` folder
- Local server-state mirroring into `data/hal-state.json` so HAL can start moving toward a shared backend model without dropping local safety
- A reusable server-side state store module so the current file mirror can later be swapped for a hosted database with less frontend churn
- Browser-uploaded calendar CSV support so `My Day` can work without depending on a server-local calendar file
- Optional Supabase-backed state storage for a free-hosting path
- A local calendar CSV sync script that can push your OneDrive-backed `calendar_export.csv` into Supabase-backed HAL state
- A Vercel-ready Express entry point and routing config so HAL can be hosted without changing the local workflow
- Teams activity-feed reminder notification scaffolding using Microsoft Graph
- SMS reminder delivery scaffolding with saved phone settings and a backend test route
- Teams-aware frontend hooks and a starter Microsoft Teams app package
- Placeholder integration area for future Microsoft Graph and Microsoft 365 features

## How to run it

HAL is designed to be iterated locally first. Microsoft integrations are optional until the core workflow feels right.

1. Run `npm install`
2. Run `npm run dev`
3. Open `http://localhost:4000/index.html`
4. Ignore Microsoft setup until you want to test those features
5. If you want AI interpretation, add `OPENAI_API_KEY` to `.env` and restart HAL
6. If you want Teams reminder testing later, add the Teams manifest ID and Microsoft scopes to `.env`, sign in with Microsoft, and restart HAL
7. If you want SMS testing later, add Twilio values to `.env` and restart HAL
8. If you want a safety snapshot before bigger changes, use the new `Backups` button in HAL to download a backup or create a local snapshot file

## Backend files included

- [package.json](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\package.json)
- [.env.example](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\.env.example)
- [docs/backend-setup.md](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\docs\backend-setup.md)
- [docs/build-log.md](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\docs\build-log.md)
- [docs/entra-app-registration.md](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\docs\entra-app-registration.md)
- [docs/online-migration-plan.md](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\docs\online-migration-plan.md)
- [docs/free-online-setup.md](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\docs\free-online-setup.md)
- [docs/calendar-csv-sync.md](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\docs\calendar-csv-sync.md)
- [docs/supabase-schema.sql](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\docs\supabase-schema.sql)
- [vercel.json](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\vercel.json)

## Teams files included

- [appPackage/manifest.json](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\appPackage\manifest.json)
- [appPackage/README.md](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\appPackage\README.md)
- [docs/teams-setup.md](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\docs\teams-setup.md)
- [hal.config.example.js](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\hal.config.example.js)

## Suggested next build steps

1. Refine the local capture and organization workflow.
2. Add a real backend database so notes sync across devices.
3. Improve document generation and export formats.
4. Add Microsoft sign-in and calendar-backed Teams meeting creation.
5. Host HAL on HTTPS and sideload it into Teams.
