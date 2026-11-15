-- ASIF AI database schema. Run in Supabase SQL Editor.
create extension if not exists pgcrypto;

create table if not exists public.conversations (
  id uuid primary key default gen_random_uuid(),
  user_id uuid not null references auth.users(id) on delete cascade,
  title text not null default 'New conversation',
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create table if not exists public.messages (
  id uuid primary key default gen_random_uuid(),
  conversation_id uuid not null references public.conversations(id) on delete cascade,
  user_id uuid not null references auth.users(id) on delete cascade,
  role text not null check (role in ('user', 'assistant')),
  content text not null,
  created_at timestamptz not null default now()
);

create index if not exists conversations_user_updated_idx on public.conversations(user_id, updated_at desc);
create index if not exists messages_conversation_created_idx on public.messages(conversation_id, created_at asc);

alter table public.conversations enable row level security;
alter table public.messages enable row level security;

create policy "Users can read their own conversations" on public.conversations
  for select to authenticated using ((select auth.uid()) = user_id);
create policy "Users can create their own conversations" on public.conversations
  for insert to authenticated with check ((select auth.uid()) = user_id);
create policy "Users can update their own conversations" on public.conversations
  for update to authenticated using ((select auth.uid()) = user_id) with check ((select auth.uid()) = user_id);
create policy "Users can delete their own conversations" on public.conversations
  for delete to authenticated using ((select auth.uid()) = user_id);

create policy "Users can read their own messages" on public.messages
  for select to authenticated using (
    (select auth.uid()) = user_id and exists (
      select 1 from public.conversations c where c.id = conversation_id and c.user_id = (select auth.uid())
    )
  );
create policy "Users can create their own messages" on public.messages
  for insert to authenticated with check (
    (select auth.uid()) = user_id and exists (
      select 1 from public.conversations c where c.id = conversation_id and c.user_id = (select auth.uid())
    )
  );
create policy "Users can delete their own messages" on public.messages
  for delete to authenticated using ((select auth.uid()) = user_id);

grant select, insert, update, delete on public.conversations to authenticated;
grant select, insert, delete on public.messages to authenticated;
