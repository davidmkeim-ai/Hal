# HAL Online Migration Plan

## Recommended first host

HAL's current shape is a single Node/Express web app with:

- a browser client
- backend API routes
- local JSON-backed state and backups
- file-based calendar CSV import

That makes a standard web-service host the cleanest first step. The best first target is **Render** because it supports:

- Node web services
- a public default URL
- custom domains later
- persistent disks for HAL's data/backups/CSV file flow

Official references:

- [Render Web Services](https://render.com/docs/web-services)
- [Render Blueprints](https://render.com/docs/infrastructure-as-code)

## What HAL will look like online

- HAL loads like a normal website
- it gets a default Render URL first
- later you can point a custom domain or subdomain to it
- the same app can later be wrapped as a desktop app if we want always-on mic behavior

## URL control

Initial URL:

- Render gives each web service its own `onrender.com` subdomain

Later:

- you can attach a custom domain like `hal.yourdomain.com`
- if the domain is a company domain, IT may need to help with DNS

## What we already prepared in code

- HAL now mirrors its state to the backend
- backend state is isolated behind a reusable store module
- storage roots are env-configurable:
  - `HAL_DATA_ROOT`
  - `HAL_BACKUP_ROOT`
  - `HAL_CSV_CALENDAR_PATH`
- a starter [render.yaml](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\render.yaml) is included

## Hosting caveat to keep in mind

HAL still uses file-based pieces today:

- `data/hal-state.json`
- `backups/*.json`
- `calendar_export.csv`

That is acceptable for an early hosted version if the host provides persistent disk storage. Longer term, we should move:

- state into a database
- backups into managed storage
- calendar import into an API or uploaded file flow

## Recommended next steps

1. Keep improving HAL locally until command understanding feels stable enough.
2. Create a git repo for HAL if you have not already.
3. Create a Render account and connect the repo.
4. Deploy using [render.yaml](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\render.yaml).
5. Confirm HAL works at the default hosted URL first.
6. Add a custom domain only after the hosted version feels stable.
7. After that, migrate from file-backed state to a proper database.
