-- Recojos (2026-10-14): OC/OPD que Friopacking recoge en el local del proveedor.
-- Compras llena la ficha (botón 🚚 en compras.html) y Almacén la ve en la pestaña
-- 📍 Recojos de inventario.html, con mapa, y marca Recogido. Una fila por documento.
-- Ejecutar una sola vez en Supabase → SQL Editor.

create table if not exists public.recojos (
  id            uuid primary key default gen_random_uuid(),
  doc           text not null unique,               -- N° de OC (0001-0012075) u OPD (0012413)
  tipo_doc      text not null default 'OC',         -- OC | OPD
  prov          text,
  ruc           text,
  idproy        text,
  proy          text,
  resp          text,
  contacto      text,                               -- persona del proveedor que entrega
  celular       text,
  direccion     text,
  distrito      text,
  referencia    text,
  maps_url      text,
  lat           double precision,
  lon           double precision,
  horario       text,                               -- horario de atención del almacén del proveedor
  f_disponible  date,                               -- desde cuándo está listo para recoger
  f_limite      date,                               -- hasta cuándo se puede recoger
  requisitos    text,                               -- EPP, SCTR, copia de OC, etc.
  obs           text,                               -- observación de Compras
  estado        text not null default 'Falta completar',  -- Falta completar | Listo para recoger | Recogido | Anulado
  f_recogido    date,
  recogido_por  text,
  obs_almacen   text,                               -- observación de Almacén
  creado_por    text,
  actualizado_por text,
  created_at    timestamptz not null default now(),
  updated_at    timestamptz not null default now()
);

-- Creada por SQL: sin este grant la plataforma muestra "permission denied for table recojos".
grant select, insert, update, delete on public.recojos to authenticated;

alter table public.recojos enable row level security;
create policy "recojos select" on public.recojos for select to authenticated using (true);
create policy "recojos insert" on public.recojos for insert to authenticated with check (true);
create policy "recojos update" on public.recojos for update to authenticated using (true) with check (true);
create policy "recojos delete" on public.recojos for delete to authenticated using (true);
