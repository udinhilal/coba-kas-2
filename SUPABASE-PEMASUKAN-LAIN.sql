-- Jalankan satu kali melalui Supabase Dashboard > SQL Editor.
-- Tabel ini menyimpan pemasukan selain iuran warga dan terpisah dari tabel expenses/payments.

create table if not exists public.other_income (
  id uuid primary key,
  owner_id uuid not null references auth.users (id) on delete cascade,
  income_date date not null default current_date,
  period_month date not null,
  title text not null,
  amount numeric(14, 0) not null check (amount > 0),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);

create index if not exists other_income_owner_date_idx
  on public.other_income (owner_id, income_date desc);

alter table public.other_income enable row level security;

grant select, insert, update, delete on public.other_income to authenticated;

drop policy if exists "Users can view their own other income" on public.other_income;
create policy "Users can view their own other income"
  on public.other_income for select
  to authenticated
  using (auth.uid() = owner_id);

drop policy if exists "Users can insert their own other income" on public.other_income;
create policy "Users can insert their own other income"
  on public.other_income for insert
  to authenticated
  with check (auth.uid() = owner_id);

drop policy if exists "Users can update their own other income" on public.other_income;
create policy "Users can update their own other income"
  on public.other_income for update
  to authenticated
  using (auth.uid() = owner_id)
  with check (auth.uid() = owner_id);

drop policy if exists "Users can delete their own other income" on public.other_income;
create policy "Users can delete their own other income"
  on public.other_income for delete
  to authenticated
  using (auth.uid() = owner_id);
