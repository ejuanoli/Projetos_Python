-- Execute no Supabase > SQL Editor

-- 1. Tabela principal
CREATE TABLE IF NOT EXISTS alertas_vbucks (
    data       DATE        PRIMARY KEY,
    enviado_em TIMESTAMPTZ DEFAULT NOW(),
    tipo       TEXT        NOT NULL CHECK (tipo IN ('vbucks', 'sem_vbucks')),
    quantidade INTEGER
);

-- 2. Habilita RLS
ALTER TABLE alertas_vbucks ENABLE ROW LEVEL SECURITY;

-- 3. Permite leitura pública (necessário para o dashboard no GitHub Pages)
CREATE POLICY "leitura_publica" ON alertas_vbucks
    FOR SELECT TO anon USING (true);

-- Nota: a escrita (INSERT/UPDATE) continua bloqueada para anon.
-- O script check_vbucks.py usa a service_role key que ignora RLS.
