-- Seguimiento de OPD (Orden de Pago) — 2026-10-14.
-- La OPD la emite Provisión/Tesorería fuera de Nisira Compras, así que no llega en el
-- Reporte Maestro: compras la registra al armar el correo (botón 📄 OPD) y le hace
-- seguimiento en la pestaña OPD de compras.html. Una fila por N° de OPD (7 dígitos).
-- Ejecutar una sola vez en Supabase → SQL Editor.

create table if not exists public.opd (
  numero        text primary key,          -- '0012413'
  fecha_opd     date,                      -- fecha impresa en la OPD
  prov          text,
  ruc           text,
  area          text,                      -- Área solicitante (ej. ALMACEN)
  idproy        text,
  proy          text,
  resp          text,                      -- responsable / supervisor que la pidió
  total         numeric,
  moneda        text,                      -- 'USD' / 'PEN'
  forma_pago    text,                      -- 'CONTADO', 'CREDITO A 30 DIAS'…
  estado        text not null default 'Por enviar',  -- Por enviar | Enviada | Entregada | Anulada
  f_envio       date,                      -- cuándo se envió al proveedor
  f_entrega     date,                      -- cuándo se recibió el pedido
  obs           text,
  creado_por    text,
  actualizado_por text,
  created_at    timestamptz not null default now(),
  updated_at    timestamptz not null default now()
);

-- Creada por SQL (no desde el Table Editor), la tabla no recibe permisos automáticos:
-- sin este grant la plataforma muestra "permission denied for table opd".
grant select, insert, update, delete on public.opd to authenticated;

alter table public.opd enable row level security;
create policy "opd select" on public.opd for select to authenticated using (true);
create policy "opd insert" on public.opd for insert to authenticated with check (true);
create policy "opd update" on public.opd for update to authenticated using (true) with check (true);
create policy "opd delete" on public.opd for delete to authenticated using (true);
