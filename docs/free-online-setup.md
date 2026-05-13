# HAL Free Online Setup

This path is designed to get HAL online without paying, while keeping your local version safe.

## Recommended Free Stack

- GitHub for source code
- Vercel Hobby for hosting the web app
- Supabase Free for HAL's database

## Step 1: Create Supabase Project

1. Go to [Supabase](https://supabase.com/).
2. Create a free project.
3. Open the SQL editor.
4. Run the script in [docs/supabase-schema.sql](C:\Users\dmk\OneDrive - RW Supply + Design\Personal Assistant App\docs\supabase-schema.sql).

## Step 2: Collect Supabase Values

From the Supabase project settings, collect:

- Project URL
- Service role key

Keep the service role key private. It belongs only in server environment variables.

## Step 3: Configure HAL

Set these environment variables locally or in Vercel:

```env
HAL_STORAGE_PROVIDER=supabase
HAL_ACCESS_PASSWORD=choose-a-strong-password
SUPABASE_URL=https://your-project.supabase.co
SUPABASE_SERVICE_ROLE_KEY=your-service-role-key
HAL_SUPABASE_STATE_KEY=default
```

`HAL_ACCESS_PASSWORD` is the cleanest way to bootstrap the first hosted login. After you sign in, HAL's Security settings can store an updated password inside HAL's backend state.

## Step 4: Deploy Hosting

Deploy HAL to Vercel from GitHub after the Supabase environment variables are set.

## Current Caveat

HAL still runs as an Express app today. The Supabase state store is now ready, but Vercel deployment may need a small serverless adapter step before the app is fully hosted there.
