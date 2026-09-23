-- ========================================================
-- TABLA DE CORTES PARA ESCUADRÓN DE COMBATE QUEST
-- ========================================================

create table if not exists quest_cortes (
    id text primary key,
    name text not null,
    date text not null,
    data jsonb not null,
    pdv_data jsonb,
    created_at timestamp with time zone default timezone('utc'::text, now()) not null
);

-- Habilitar RLS (Seguridad a nivel de fila)
alter table quest_cortes enable row level security;

-- Política de lectura pública (para todos los espectadores/tiendas/líderes)
create policy "Lectura publica de cortes"
    on quest_cortes for select
    to anon, authenticated
    using (true);

-- Política de gestión de cortes (inserción, actualización y borrado)
create policy "Permitir guardar y gestionar cortes"
    on quest_cortes for all
    to anon, authenticated
    using (true)
    with check (true);

-- Habilitar replicación Realtime
alter publication supabase_realtime add table quest_cortes;
