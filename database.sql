-- Run in Supabase SQL Editor. Cloud records are private to their authenticated owner.
create table if not exists public.user_state (
 user_id uuid primary key references auth.users(id) on delete cascade,
 data jsonb not null default '{"trips":[],"items":[]}'::jsonb,
 updated_at timestamptz not null default now(),
 constraint data_shape check (jsonb_typeof(data->'trips')='array' and jsonb_typeof(data->'items')='array')
);
alter table public.user_state enable row level security;
revoke all on public.user_state from anon;
revoke all on public.user_state from authenticated;
grant select,insert,update on public.user_state to authenticated;
drop policy if exists "Owner can read" on public.user_state;
drop policy if exists "Owner can insert" on public.user_state;
drop policy if exists "Owner can update" on public.user_state;
create policy "Owner can read" on public.user_state for select to authenticated using ((select auth.uid())=user_id);
create policy "Owner can insert" on public.user_state for insert to authenticated with check ((select auth.uid())=user_id);
create policy "Owner can update" on public.user_state for update to authenticated using ((select auth.uid())=user_id) with check ((select auth.uid())=user_id);
-- Prepared metadata model: NO access tokens are stored in the browser or these tables.
create table if not exists public.provider_connections (
 id uuid primary key default gen_random_uuid(),
 user_id uuid not null references auth.users(id) on delete cascade,
 provider text not null check (provider in ('microsoft','booking','airbnb','other')),
 provider_account_id text,
 status text not null default 'not_connected' check(status in ('not_connected','connected','expired','revoked')),
 last_synced_at timestamptz,
 created_at timestamptz not null default now(),
 unique(user_id,provider,provider_account_id)
);
alter table public.provider_connections enable row level security;
revoke all on public.provider_connections from anon,authenticated;
grant select on public.provider_connections to authenticated;
create policy "Connections owner read" on public.provider_connections for select to authenticated using ((select auth.uid())=user_id);
-- Future: trusted backend/service role manages connection status and encrypted OAuth tokens.
create table if not exists public.imported_messages (
 id uuid primary key default gen_random_uuid(),
 user_id uuid not null references auth.users(id) on delete cascade,
 provider text not null,
 provider_message_id text not null,
 booking_type text,
 parsed_booking jsonb,
 imported_at timestamptz not null default now(),
 unique(user_id,provider,provider_message_id)
);
alter table public.imported_messages enable row level security;
revoke all on public.imported_messages from anon,authenticated;
grant select on public.imported_messages to authenticated;
create policy "Imported messages owner read" on public.imported_messages for select to authenticated using ((select auth.uid())=user_id);
