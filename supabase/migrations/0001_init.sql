-- 0001_init.sql — deck/slide/revision schema with row-level security.
-- Spec fragments live HERE, not in PowerPoint slide tags. Tags carry only a
-- slide_uuid pointing at rows in this schema (tags have size limits and no
-- history; Postgres gives us both).

create extension if not exists "pgcrypto";

create table public.decks (
  id          uuid primary key default gen_random_uuid(),
  owner       uuid not null references auth.users (id) on delete cascade,
  title       text not null default 'Untitled deck',
  language    text,                       -- BCP-47; null = auto-detect
  audience    text,
  created_at  timestamptz not null default now(),
  updated_at  timestamptz not null default now()
);

create table public.slides (
  id          uuid primary key default gen_random_uuid(),
  deck_id     uuid not null references public.decks (id) on delete cascade,
  position    int  not null,              -- 0-based order within the deck
  spec        jsonb not null,             -- one element of deck_spec["slides"]
  gate_result jsonb,                      -- last per-slide gate output
  detached    boolean not null default false,  -- user opted out of regeneration
  created_at  timestamptz not null default now(),
  updated_at  timestamptz not null default now(),
  unique (deck_id, position) deferrable initially deferred
);

-- Full history of every spec change — powers spec-level undo/rollback,
-- which PowerPoint-side undo cannot give us (insert+delete = 2 undo steps).
create table public.revisions (
  id          bigint generated always as identity primary key,
  slide_id    uuid not null references public.slides (id) on delete cascade,
  spec        jsonb not null,
  source      text not null check (source in ('generate','chat_edit','manual','rollback')),
  prompt      text,                       -- the chat command that caused this, if any
  created_at  timestamptz not null default now()
);

create index on public.slides (deck_id, position);
create index on public.revisions (slide_id, created_at desc);

-- updated_at maintenance
create or replace function public.touch_updated_at()
returns trigger language plpgsql as $$
begin new.updated_at = now(); return new; end $$;

create trigger decks_touch  before update on public.decks
  for each row execute function public.touch_updated_at();
create trigger slides_touch before update on public.slides
  for each row execute function public.touch_updated_at();

-- Row Level Security: owner-only for everything.
alter table public.decks     enable row level security;
alter table public.slides    enable row level security;
alter table public.revisions enable row level security;

create policy decks_owner on public.decks
  for all using (owner = auth.uid()) with check (owner = auth.uid());

create policy slides_owner on public.slides
  for all using (exists (select 1 from public.decks d
                         where d.id = deck_id and d.owner = auth.uid()))
  with check   (exists (select 1 from public.decks d
                         where d.id = deck_id and d.owner = auth.uid()));

create policy revisions_owner on public.revisions
  for all using (exists (select 1 from public.slides s
                         join public.decks d on d.id = s.deck_id
                         where s.id = slide_id and d.owner = auth.uid()))
  with check   (exists (select 1 from public.slides s
                         join public.decks d on d.id = s.deck_id
                         where s.id = slide_id and d.owner = auth.uid()));

-- Storage bucket for rendered artifacts (create via dashboard or CLI):
--   supabase storage buckets create renders --public=false
