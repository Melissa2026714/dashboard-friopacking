-- Fase 3: detalle de ítems por OC (SKUS / SKUS_P4) — de data.json a Supabase.
-- Correr una sola vez en el SQL Editor de Supabase (proyecto gxjwhubmwsizrxbmavlz).

create table if not exists public.oc_items (
  id bigserial primary key,
  oc text not null,
  cod text,
  req text not null default '',
  prod text,
  cant_ord numeric,
  cant_rec numeric,
  cant_pend numeric,
  foc text,
  fent text,
  estado text,
  ucomp text,
  prov text,
  resp text,
  idproy text,
  proy text,
  freq text,
  frec text,
  moneda text,
  punit numeric,
  tc numeric,
  unid text,
  unique (oc, cod, req)
);
create index if not exists oc_items_oc_idx on public.oc_items (oc);

alter table public.oc_items enable row level security;
create policy "oc_items select" on public.oc_items for select to authenticated using (true);
create policy "oc_items insert" on public.oc_items for insert to authenticated with check (true);
create policy "oc_items update" on public.oc_items for update to authenticated using (true) with check (true);
create policy "oc_items delete" on public.oc_items for delete to authenticated using (true);

create table if not exists public.p4_items (
  id bigserial primary key,
  oc text not null,
  cod text,
  ped text not null default '',
  prod text,
  cant_ord numeric,
  cant_rec numeric,
  cant_pend numeric,
  foc text,
  fent text,
  estado text,
  ucomp text,
  prov text,
  resp text,
  frec text,
  unid text,
  unique (oc, cod, ped)
);
create index if not exists p4_items_oc_idx on public.p4_items (oc);

alter table public.p4_items enable row level security;
create policy "p4_items select" on public.p4_items for select to authenticated using (true);
create policy "p4_items insert" on public.p4_items for insert to authenticated with check (true);
create policy "p4_items update" on public.p4_items for update to authenticated using (true) with check (true);
create policy "p4_items delete" on public.p4_items for delete to authenticated using (true);
