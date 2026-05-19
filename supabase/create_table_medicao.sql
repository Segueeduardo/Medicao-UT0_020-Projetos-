-- =============================================================
-- TABELA: medicao_obras_projetos
-- Origem: MEDICÃO_OBR_PRJ.xlsx | Aba: sisteplant
-- Criado: 2026-05-18
-- =============================================================

CREATE TABLE IF NOT EXISTS medicao_obras_projetos (
  id                               BIGSERIAL PRIMARY KEY,

  -- Identificação da OS
  numero_os                        INTEGER NOT NULL,

  -- Dados do técnico
  tecnico                          VARCHAR(20),
  denominacao_tecnico              VARCHAR(255),
  denominacao_especialidade        VARCHAR(255),
  denominacao_oficina              VARCHAR(255),

  -- Tipo de serviço e unidade
  denominacao_tipo_solicitacao     VARCHAR(255),
  denominacao_unidade_negocio      VARCHAR(255),

  -- Datas e horários do serviço
  data_inicio_trabalho             TIMESTAMP WITH TIME ZONE,
  data_final_trabalho              TIMESTAMP WITH TIME ZONE,

  -- Tempo e valores financeiros
  tempo_trabalho_feedback_mao_obra INTERVAL,
  preco_hora                       NUMERIC(10, 4),
  valor_mao_de_obra                NUMERIC(12, 2),

  -- Controle de registro
  created_at                       TIMESTAMP WITH TIME ZONE DEFAULT NOW()
);

-- =============================================================
-- ÍNDICES PARA PERFORMANCE
-- =============================================================
CREATE INDEX IF NOT EXISTS idx_medicao_numero_os
  ON medicao_obras_projetos(numero_os);

CREATE INDEX IF NOT EXISTS idx_medicao_tecnico
  ON medicao_obras_projetos(tecnico);

CREATE INDEX IF NOT EXISTS idx_medicao_data_inicio
  ON medicao_obras_projetos(data_inicio_trabalho);

CREATE INDEX IF NOT EXISTS idx_medicao_unidade_negocio
  ON medicao_obras_projetos(denominacao_unidade_negocio);

CREATE INDEX IF NOT EXISTS idx_medicao_oficina
  ON medicao_obras_projetos(denominacao_oficina);

-- =============================================================
-- ROW LEVEL SECURITY (RLS)
-- =============================================================
ALTER TABLE medicao_obras_projetos ENABLE ROW LEVEL SECURITY;

-- Qualquer usuário autenticado pode ler
CREATE POLICY "Leitura autenticada"
  ON medicao_obras_projetos
  FOR SELECT
  TO authenticated
  USING (true);

-- Apenas service_role pode inserir/atualizar/deletar (importação via backend)
CREATE POLICY "Escrita service_role"
  ON medicao_obras_projetos
  FOR ALL
  TO service_role
  USING (true)
  WITH CHECK (true);

-- =============================================================
-- COMENTÁRIOS NAS COLUNAS (documentação)
-- =============================================================
COMMENT ON TABLE medicao_obras_projetos IS
  'Medição de mão de obra por OS, importada do Sisteplant';

COMMENT ON COLUMN medicao_obras_projetos.numero_os IS
  'Número da Ordem de Serviço';

COMMENT ON COLUMN medicao_obras_projetos.tecnico IS
  'Código do técnico responsável';

COMMENT ON COLUMN medicao_obras_projetos.preco_hora IS
  'Preço por hora em R$';

COMMENT ON COLUMN medicao_obras_projetos.valor_mao_de_obra IS
  'Valor total de mão de obra em R$';

COMMENT ON COLUMN medicao_obras_projetos.tempo_trabalho_feedback_mao_obra IS
  'Tempo de trabalho registrado (ex: 4h = 14400s)';
