create table if not exists public.hal_state (
  id text primary key,
  state jsonb not null,
  updated_at timestamptz not null default now()
);

alter table public.hal_state enable row level security;

-- HAL writes through the server using the Supabase service role key.
-- Do not expose the service role key in browser code.
create policy "No direct browser access to HAL state"
on public.hal_state
for all
using (false)
with check (false);
