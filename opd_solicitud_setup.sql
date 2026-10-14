-- Seguimiento de OPD, v2 (2026-10-14): agrega la etapa de SOLICITUD.
-- Flujo: Solicitada (correo a Finanzas con la cotización) → Aprobada (Sr. Álvaro / Anthony)
-- → OPD recibida (Provisión emitió la OPD) → Enviada (al proveedor) → Entregada · Anulada.
-- Una solicitud todavía no tiene N° de OPD, así que la llave pasa de "numero" a "id"
-- (numero queda único pero opcional). No borra ni modifica ningún registro existente,
-- salvo renombrar el estado antiguo "Por enviar" a "OPD recibida".
-- Ejecutar una sola vez en Supabase → SQL Editor, DESPUÉS de opd_setup.sql.

alter table public.opd add column if not exists id uuid not null default gen_random_uuid();
alter table public.opd add column if not exists codigo text;          -- SOL-OPD-2026-001
alter table public.opd add column if not exists descripcion text;     -- qué se compra
alter table public.opd add column if not exists cotizacion text;      -- nombre del archivo de cotización
alter table public.opd add column if not exists f_solicitud date;
alter table public.opd add column if not exists f_aprobacion date;
alter table public.opd add column if not exists f_opd_recibida date;

alter table public.opd drop constraint if exists opd_pkey;
alter table public.opd add primary key (id);
alter table public.opd alter column numero drop not null;
create unique index if not exists opd_numero_key on public.opd (numero);
create unique index if not exists opd_codigo_key on public.opd (codigo);

grant select, insert, update, delete on public.opd to authenticated;

alter table public.opd alter column estado set default 'Solicitada';
update public.opd set estado = 'OPD recibida' where estado = 'Por enviar';
