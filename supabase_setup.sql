-- Execute esse SQL no Supabase > SQL Editor uma única vez

CREATE TABLE IF NOT EXISTS alertas_vbucks (
    data       DATE        PRIMARY KEY,
    enviado_em TIMESTAMPTZ DEFAULT NOW(),
    tipo       TEXT        NOT NULL CHECK (tipo IN ('vbucks', 'sem_vbucks')),
    quantidade INTEGER
);

-- Impede qualquer outro usuário de ler/alterar (só o service_role acessa)
ALTER TABLE alertas_vbucks ENABLE ROW LEVEL SECURITY;
