-- Fase 3: detalle de ítems por OC (SKUS / SKUS_P4) — de data.json a Supabase.
-- v2 (2026-08-27): la llave (oc,cod,req)/(oc,cod,ped) NO es única — el Maestro trae
-- varias filas para el mismo producto/requerimiento dentro de una OC (recepciones
-- parciales), y eso rompía el upsert ("ON CONFLICT command cannot affect row a
-- second time"). Se reemplaza por (oc, seq), donde seq es la posición de cada fila
-- dentro de su OC — siempre única, no se pierde ninguna línea de movimiento.
-- Las tablas quedaron vacías (todos los intentos de guardar fallaron), así que se
-- recrean limpias en vez de alterar la v1.

drop table if exists public.oc_items;
drop table if exists public.p4_items;

create table public.oc_items (
  id bigserial primary key,
  oc text not null,
  seq integer not null,
  cod text,
  req text,
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
  unique (oc, seq)
);
create index oc_items_oc_idx on public.oc_items (oc);

alter table public.oc_items enable row level security;
create policy "oc_items select" on public.oc_items for select to authenticated using (true);
create policy "oc_items insert" on public.oc_items for insert to authenticated with check (true);
create policy "oc_items update" on public.oc_items for update to authenticated using (true) with check (true);
create policy "oc_items delete" on public.oc_items for delete to authenticated using (true);

create table public.p4_items (
  id bigserial primary key,
  oc text not null,
  seq integer not null,
  cod text,
  ped text,
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
  unique (oc, seq)
);
create index p4_items_oc_idx on public.p4_items (oc);

alter table public.p4_items enable row level security;
create policy "p4_items select" on public.p4_items for select to authenticated using (true);
create policy "p4_items insert" on public.p4_items for insert to authenticated with check (true);
create policy "p4_items update" on public.p4_items for update to authenticated using (true) with check (true);
create policy "p4_items delete" on public.p4_items for delete to authenticated using (true);
