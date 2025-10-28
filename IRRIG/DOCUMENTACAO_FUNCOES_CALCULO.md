# DOCUMENTAÇÃO - FUNÇÕES DE CÁLCULO DO SISTEMA DE ORÇAMENTO
**Sistema de Irrigação Penápolis**
**Módulo:** Orcament
**Data da Documentação:** 10/10/2025

---

## ÍNDICE
1. [Visão Geral](#visão-geral)
2. [Funções Processa](#funções-processa)
3. [Funções AtualizaValores](#funções-atualizavalores)
4. [Funções Distribui](#funções-distribui)
5. [Funções SuperAtualiza](#funções-superatualiza)
6. [Rotinas de Reprocessamento em Lote](#rotinas-de-reprocessamento-em-lote)
7. [Função AjustaValores](#função-ajustavalores)
8. [Função CalculaImposto](#função-calculaimposto)
9. [Fluxo de Execução](#fluxo-de-execução)
10. [Regras de Usuário](#regras-de-usuário)

---

## VISÃO GERAL

O sistema de orçamento possui **4 tipos principais de funções de cálculo**:

| Tipo de Função | Quando Executa | Escopo | Recalcula Impostos? |
|----------------|----------------|--------|---------------------|
| **Processa**   | Ao editar item no grid | 1 item por vez | ✅ SIM (exceto YGOR/JUCELI/MAYSA) |
| **AtualizaValores** | Após ProcessaXXX | Todos os itens | ⚠️ Distribui desconto/frete |
| **Distribui*** | Ao salvar desconto/frete global ou quando chamado explicitamente | Todos os itens | ❌ Apenas rateia financeiro |
| **SuperAtualiza** | Tecla F2 (manual) | Todos os itens | ✅ SIM (força recálculo) |
| **AjustaValores** | Após editar cabeçalho | Calcula totais | ❌ NÃO recalcula grids |

*As funções `DistribuiDescontoTotal` e `DistribuiFreteTotal` são chamadas pelas rotinas `AtualizaValores` e também diretamente quando o usuário altera os campos financeiros do cabeçalho.

---

## FUNÇÕES PROCESSA

### 📌 **ProcessaProdutos** (Linha 669)

**Responsabilidade:**
Calcula impostos e valores fiscais de **UM produto específico** quando ele é editado/incluído no grid.

**Quando é chamada:**
- Ao **incluir** um produto no grid
- Ao **editar** quantidade, valor unitário ou dados do produto
- Via evento `Grid_AfterUpdateRecord` → `ExecutaGrid` → `PROCESSOS_DIRETOS`

**O que faz:**
1. **Atualiza Valor Total:** `Valor Total = Quantidade × Valor Unitário`
2. **Se Orçamento NÃO for Avulso E usuário NÃO for YGOR/JUCELI/MAYSA:**
   - Calcula **CST** (Código de Situação Tributária)
   - Calcula **CFOP** (Código Fiscal de Operação)
   - Calcula **Base de Cálculo ICMS**
   - Calcula **Valor ICMS**
   - Calcula **Alíquota ICMS**
   - Calcula **Percentual de Redução**
   - Calcula **Valor IPI**
   - Calcula **Alíquota IPI**
   - Calcula **Diferido**
   - Calcula **PIS** (com redução de 48,1% para NCM específicos: 84248, 7309, 87162000)
   - Calcula **COFINS** (com redução de 48,1% para NCM específicos)
   - Calcula **IVA** (Índice de Valor Agregado)
   - Calcula **Base de Cálculo ST** (Substituição Tributária)
   - Calcula **Valor ICMS ST**
   - Calcula **Alíquota ICMS ST**
   - Calcula **Valor do Tributo** (soma de todos os impostos)

**Entrega Futura:**
- Se `Entrega_Futura = True`:
  - Define CFOP 5922 (SP) ou 6922 (outros estados)
  - Define CST 90
  - Zera ICMS, Base de Cálculo, Alíquotas

**Redução PIS/COFINS (48,1%):**
- Aplica redução de 48,1% na base de cálculo para produtos fabricados (NCM 84248*, 7309*, 87162000)
- Produtos adquiridos de terceiros **NÃO** têm redução

**Alíquotas PIS/COFINS:**
- **Com redução (produtos fabricados):**
  - PIS: 2%
  - COFINS: 9,6%
- **Sem redução (produtos adquiridos):**
  - PIS: 1,65%
  - COFINS: 7,6%

**Chama ao final (somente para usuários comuns):**
- `AtualizaValoresProdutos` (distribui desconto/frete proporcionalmente)

**Retorno:**
- `Boolean` — `True` confirma que a gravação ocorreu sem erro; `False` sinaliza falha e evita avanços na rotina chamadora.

**Banco de dados:**
- Tabela: `Produtos do Orçamento`
- Atualiza: Todos os campos fiscais e tributários

---

### 📌 **ProcessaPecas** (Linha 917)

**Responsabilidade:**
Calcula impostos e valores fiscais de **UMA peça específica** quando ela é editada/incluída no grid.

**Quando é chamada:**
- Ao **incluir** uma peça no grid
- Ao **editar** quantidade, valor unitário ou dados da peça
- Via evento `Grid_AfterUpdateRecord` → `ExecutaGrid` → `PROCESSOS_DIRETOS`

**O que faz:**
1. **Atualiza Valor Total:** `Valor Total = Quantidade × Valor Unitário`
2. **Se Orçamento NÃO for Avulso E usuário NÃO for YGOR/JUCELI/MAYSA:**
   - Calcula **CST**
   - Calcula **CFOP**
   - Calcula **Base de Cálculo ICMS**
   - Calcula **Valor ICMS**
   - Calcula **Alíquota ICMS**
   - Calcula **Percentual de Redução**
   - Calcula **Valor IPI**
   - Calcula **Alíquota IPI**
   - Calcula **Diferido**
   - Calcula **IVA**
   - Calcula **Base de Cálculo ST**
   - Calcula **Valor ICMS ST**
   - Calcula **Alíquota ICMS ST**
   - Calcula **PIS** (com redução de 48,1% se desconto/frete existir)
   - Calcula **COFINS** (com redução de 48,1% se desconto/frete existir)
   - Calcula **Valor do Tributo**

**Entrega Futura:**
- Se `Entrega_Futura = True`:
  - Define CFOP 5922 (SP) ou 6922 (outros estados)
  - Define CST 90
  - Zera ICMS, Base de Cálculo, Alíquotas
  - PIS/COFINS calculados posteriormente via `AtualizaValoresPecas`

**Cálculo PIS/COFINS (Peças):**
- **Base de cálculo:**
  - **COM desconto/frete:** `(Valor + Frete - Desconto - ICMS) - 48,1%`
  - **SEM desconto/frete:** `(Valor - ICMS) - 48,1%`
- **Alíquotas fixas:**
  - PIS: **2%**
  - COFINS: **9,6%**

**Chama ao final (somente para usuários comuns):**
- `AtualizaValoresPecas` (distribui desconto/frete proporcionalmente)

**Retorno:**
- `Boolean` — segue o mesmo padrão de sucesso/falha utilizado em `ProcessaProdutos`.

**Banco de dados:**
- Tabela: `Peças do Orçamento`
- Atualiza: Todos os campos fiscais e tributários

---

### 📌 **ProcessaConjuntos** (Linha 783)

**Responsabilidade:**
Calcula impostos e valores fiscais de **UM conjunto específico** quando ele é editado/incluído no grid.

**Quando é chamada:**
- Ao **incluir** um conjunto no grid
- Ao **editar** quantidade, valor unitário ou dados do conjunto
- Via evento `Grid_AfterUpdateRecord` → `ExecutaGrid` → `PROCESSOS_DIRETOS`

**O que faz:**
1. **Atualiza Valor Total:** `Valor Total = Quantidade × Valor Unitário`
2. **Se Orçamento NÃO for Avulso E usuário NÃO for YGOR/JUCELI/MAYSA:**
   - Calcula **CST**
   - Calcula **CFOP**
   - Calcula **Base de Cálculo ICMS**
   - Calcula **Valor ICMS**
   - Calcula **Valor IPI**
   - Calcula **Alíquota ICMS**
   - Calcula **Alíquota IPI**
   - Calcula **Percentual de Redução**
   - Calcula **IVA**
   - Calcula **Base de Cálculo ST**
   - Calcula **Valor ICMS ST**
   - Calcula **Alíquota ICMS ST**
   - Calcula **Diferido**
   - Calcula **PIS** (base = Valor - ICMS)
   - Calcula **COFINS** (base = Valor - ICMS)
   - Calcula **Valor do Tributo**

**Entrega Futura:**
- Se `Entrega_Futura = True`:
  - Define CFOP 5922 (SP) ou 6922 (outros estados)
  - Define CST 90
  - Zera ICMS, IPI, Base de Cálculo, Alíquotas, IVA, ST

**Garantia de Valores Positivos:**
- Executa UPDATE para garantir que PIS, COFINS, Tributos, Bc pis, Bc cofins **nunca sejam negativos**

**Chama ao final (somente para usuários comuns):**
- `AtualizaValoresConjuntos` (distribui desconto/frete proporcionalmente)

**Retorno:**
- `Boolean` — indica à rotina chamadora se o cálculo concluiu corretamente.

**Banco de dados:**
- Tabela: `Conjuntos do Orçamento`
- Atualiza: Todos os campos fiscais e tributários

---

### 📌 **ProcessaServicos** (Linha 877)

**Responsabilidade:**
Calcula apenas o **Valor Total** de um serviço (serviços não têm impostos de produto).

**Quando é chamada:**
- Ao editar quantidade ou valor unitário de um serviço

**O que faz:**
1. **Atualiza Valor Total:** `Valor Total = Quantidade × Valor Unitário`

**Observação:**
- Serviços **NÃO têm cálculo de impostos** (ICMS, IPI, PIS, COFINS)
- Apenas atualiza valor total

**Banco de dados:**
- Tabela: `Serviços do Orçamento`
- Atualiza: Apenas `Valor Total`

**Retorno:**
- `Boolean` — `True` quando o valor total é gravado no registro; `False` somente em caso de erro de gravação.

---

## FUNÇÕES ATUALIZAVALORES

### 📌 **AtualizaValoresProdutos** (Linha 4261)

**Responsabilidade:**
Distribui **desconto e frete proporcionalmente** entre TODOS os produtos do orçamento e recalcula PIS/COFINS com base distribuída.

**Quando é chamada:**
- Após `ProcessaProdutos` (somente usuários comuns)
- **NÃO** é chamada mais por `AjustaValores`

**O que faz:**

1. **Calcula Totais Gerais:**
   - Soma valor total de produtos
   - Soma valor total de conjuntos
   - Soma valor total de peças
   - Total Geral = Produtos + Conjuntos + Peças

2. **Calcula Desconto e Frete Proporcionais:**
   - `% Produtos = Valor Produtos ÷ Total Geral`
   - `Desconto Proporcional Produtos = Desconto Orçamento × % Produtos`
   - `Frete Proporcional Produtos = Frete Orçamento × % Produtos`

3. **Para cada produto:**
   - `% do Produto = Valor do Produto ÷ Total Produtos`
   - `Desconto do Produto = Desconto Proporcional × % do Produto`
   - `Frete do Produto = Frete Proporcional × % do Produto`
   - Atualiza campos `Valor do Desconto` e `Valor do Frete` no banco

4. **Recalcula PIS/COFINS considerando desconto/frete:**
   - **Base:** `(Valor + Frete - Desconto - ICMS)`
   - Aplica redução de 48,1% se NCM específico
   - Recalcula Valor PIS e Valor COFINS

**Importante:**
- **NÃO recalcula impostos básicos** (ICMS, IPI, ST)
- Apenas **distribui financeiro** (desconto/frete)
- Recalcula **PIS/COFINS** com base ajustada

**Banco de dados:**
- Tabela: `Produtos do Orçamento`
- Atualiza: `Valor do Desconto`, `Valor do Frete`, `Valor do PIS`, `Valor do Cofins`, `Bc pis`, `Bc cofins`

---

### 📌 **AtualizaValoresPecas** (Linha 3868)

**Responsabilidade:**
Distribui **desconto e frete proporcionalmente** entre TODAS as peças do orçamento e recalcula PIS/COFINS com base distribuída.

**Quando é chamada:**
- Após `ProcessaPecas` (somente usuários comuns)
- **NÃO** é chamada mais por `AjustaValores`

**O que faz:**

1. **Calcula Totais Gerais:**
   - Soma valor total de produtos
   - Soma valor total de conjuntos
   - Soma valor total de peças
   - Total Geral = Produtos + Conjuntos + Peças

2. **Calcula Desconto e Frete Proporcionais:**
   - `% Peças = Valor Peças ÷ Total Geral`
   - `Desconto Proporcional Peças = Desconto Orçamento × % Peças`
   - `Frete Proporcional Peças = Frete Orçamento × % Peças`

3. **Para cada peça:**
   - `% da Peça = Valor da Peça ÷ Total Peças`
   - `Desconto da Peça = Desconto Proporcional × % da Peça`
   - `Frete da Peça = Frete Proporcional × % da Peça`
   - Atualiza campos `Valor do Desconto` e `Valor do Frete` no banco

4. **Recalcula PIS/COFINS considerando desconto/frete:**
   - **Base:** `(Valor + Frete - Desconto - ICMS) - 48,1%`
   - **Alíquotas fixas:**
     - PIS: 2%
     - COFINS: 9,6%
   - Recalcula Valor PIS e Valor COFINS

5. **Tratamento Entrega Futura:**
   - Se `Entrega_Futura = True` e `UF <> SP`:
     - Aplica **redução regional** na base PIS/COFINS
     - Redução varia por UF (consulta tabela `Redução Regional`)

**Importante:**
- **NÃO recalcula impostos básicos** (ICMS, IPI, ST)
- Apenas **distribui financeiro** (desconto/frete)
- Recalcula **PIS/COFINS** com base ajustada
- **Sempre aplica redução de 48,1%** para peças

**Banco de dados:**
- Tabela: `Peças do Orçamento`
- Atualiza: `Valor do Desconto`, `Valor do Frete`, `Valor do PIS`, `Valor do Cofins`, `Bc pis`, `Bc cofins`, `Aliq do pis`, `Aliq do cofins`

---

### 📌 **AtualizaValoresConjuntos** (Linha 3658)

**Responsabilidade:**
Distribui **desconto e frete proporcionalmente** entre TODOS os conjuntos do orçamento e recalcula PIS/COFINS com base distribuída.

**Quando é chamada:**
- Após `ProcessaConjuntos` (somente usuários comuns)
- **NÃO** é chamada mais por `AjustaValores`

**O que faz:**

1. **Calcula Totais Gerais:**
   - Soma valor total de produtos
   - Soma valor total de conjuntos
   - Soma valor total de peças
   - Total Geral = Produtos + Conjuntos + Peças

2. **Calcula Desconto e Frete Proporcionais:**
   - `% Conjuntos = Valor Conjuntos ÷ Total Geral`
   - `Desconto Proporcional Conjuntos = Desconto Orçamento × % Conjuntos`
   - `Frete Proporcional Conjuntos = Frete Orçamento × % Conjuntos`

3. **Para cada conjunto:**
   - `% do Conjunto = Valor do Conjunto ÷ Total Conjuntos`
   - `Desconto do Conjunto = Desconto Proporcional × % do Conjunto`
   - `Frete do Conjunto = Frete Proporcional × % do Conjunto`
   - Atualiza campos `Valor do Desconto` e `Valor do Frete` no banco

4. **Recalcula PIS/COFINS considerando desconto/frete:**
   - **Base:** `(Valor + Frete - Desconto - ICMS)`
   - Recalcula Valor PIS e Valor COFINS

**Importante:**
- **NÃO recalcula impostos básicos** (ICMS, IPI, ST)
- Apenas **distribui financeiro** (desconto/frete)
- Recalcula **PIS/COFINS** com base ajustada

**Banco de dados:**
- Tabela: `Conjuntos do Orçamento`
- Atualiza: `Valor do Desconto`, `Valor do Frete`, `Valor do PIS`, `Valor do Cofins`, `Bc pis`, `Bc cofins`

---

## FUNÇÕES DISTRIBUI

### 📌 **DistribuiDescontoTotal** (Linha 2974)

**Responsabilidade:**
Rateia o **desconto global** informado no cabeçalho entre todos os itens (produtos, conjuntos e peças) proporcionalmente ao valor bruto de cada um.

**Quando é chamada:**
- Automaticamente por `AtualizaValoresProdutos`, `AtualizaValoresConjuntos` e `AtualizaValoresPecas`.
- Diretamente quando o usuário altera os campos de desconto no cabeçalho financeiro (`Index = 32`).
- Após o recálculo completo (`RecalcularImpostosTodos`) para alinhar os itens ao desconto vigente.

**O que faz:**
1. Calcula o valor bruto de cada item (Quantidade × Valor Unitário).
2. Soma todos os valores brutos para obter o denominador do rateio.
3. Calcula a fração de desconto de cada item e grava em `[Valor Do Desconto]`.
4. Ajusta o último item processado para eliminar diferenças de arredondamento.
5. Atualiza os grids visuais (`Grid(0)`, `Grid(1)`, `Grid(3)`).

**Retorno / Observações:**
- Não devolve valor; atua diretamente nos registros do banco.
- Se o desconto global for zero, encerra imediatamente sem tocar nos itens.

### 📌 **DistribuiFreteTotal** (Linha 6715)

**Responsabilidade:**
Rateia o **frete global** do orçamento entre todos os itens proporcionais ao valor bruto, garantindo que os campos de frete item a item reflitam o cabeçalho.

**Quando é chamada:**
- Pelas rotinas `AtualizaValoresProdutos`, `AtualizaValoresConjuntos` e `AtualizaValoresPecas` logo após o rateio de desconto.
- Ao editar o campo de frete no cabeçalho (`Index = 45`).
- Durante rotinas de recálculo completo (`RecalcularImpostosTodos`).

**O que faz:**
1. Zera todos os fretes individuais quando o frete global é 0.
2. Caso contrário, percorre todas as tabelas de itens, calculando o rateio proporcional.
3. Usa `ContaRegs` para saber quantos registros existem e ajustar o último item ao centavo.
4. Rebind dos grids para refletir o novo frete.

**Retorno / Observações:**
- Não devolve valor; persiste alterações diretamente nas tabelas.
- Mantém consistência entre financeiro do cabeçalho e valores utilizados em PIS/COFINS.

**Funções de Apoio:**
- `ContaRegs` (Linha 3107) — retorna o número de registros por tabela e é usada para ajustar rateios sem acumular erro.

---

## FUNÇÕES SUPERATUALIZA

### 📌 **SuperAtualizaProdutos** (Linha 2016)

**Responsabilidade:**
**Recalcula TODOS os impostos** de TODOS os produtos do orçamento (força recálculo completo).

**Quando é chamada:**
- Usuário pressiona **F2** manualmente no grid de produtos
- Botão/menu "Recalcular Impostos"

**O que faz:**

1. **Para CADA produto do orçamento:**
   - Busca dados do produto (NCM, tipo, classificação fiscal)
   - Recalcula **CST**
   - Recalcula **CFOP**
   - Recalcula **Base de Cálculo ICMS**
   - Recalcula **Valor ICMS**
   - Recalcula **Alíquota ICMS**
   - Recalcula **Percentual de Redução**
   - Recalcula **Valor IPI**
   - Recalcula **Alíquota IPI**
   - Recalcula **Diferido**
   - Recalcula **PIS** (com redução 48,1% se aplicável)
   - Recalcula **COFINS** (com redução 48,1% se aplicável)
   - Recalcula **IVA**
   - Recalcula **Base de Cálculo ST**
   - Recalcula **Valor ICMS ST**
   - Recalcula **Alíquota ICMS ST**
   - Recalcula **Valor do Tributo**

2. **Tratamento Entrega Futura:**
   - Aplica lógica específica para entrega futura (CFOP 5922/6922, CST 90)

3. **Atualiza valores no banco de dados**

**Chama ao final:**
- `AtualizaValoresProdutos` (distribui desconto/frete)
- `AjustaValores` (atualiza totais do orçamento)

**Importante:**
- **RECALCULA TUDO** (sobrescreve valores editados manualmente)
- **NÃO respeita** edições manuais de YGOR/JUCELI/MAYSA
- Usar apenas quando necessário forçar recálculo total

**Banco de dados:**
- Tabela: `Produtos do Orçamento`
- Atualiza: **TODOS** os campos fiscais e tributários

---

### 📌 **SuperAtualizaConjuntos** (Linha 2082)

**Responsabilidade:**
**Recalcula TODOS os impostos** de TODOS os conjuntos do orçamento (força recálculo completo).

**Quando é chamada:**
- Usuário pressiona **F2** manualmente no grid de conjuntos
- Botão/menu "Recalcular Impostos"

**O que faz:**

1. **Para CADA conjunto do orçamento:**
   - Busca dados do conjunto
   - Recalcula **CST**
   - Recalcula **CFOP**
   - Recalcula **Base de Cálculo ICMS**
   - Recalcula **Valor ICMS**
   - Recalcula **Alíquota ICMS**
   - Recalcula **Percentual de Redução**
   - Recalcula **Valor IPI**
   - Recalcula **Alíquota IPI**
   - Recalcula **Diferido**
   - Recalcula **PIS**
   - Recalcula **COFINS**
   - Recalcula **IVA**
   - Recalcula **Base de Cálculo ST**
   - Recalcula **Valor ICMS ST**
   - Recalcula **Alíquota ICMS ST**
   - Recalcula **Valor do Tributo**

2. **Tratamento Entrega Futura:**
   - Aplica lógica específica para entrega futura

3. **Atualiza valores no banco de dados**

**Chama ao final:**
- `AtualizaValoresConjuntos` (distribui desconto/frete)
- `AjustaValores` (atualiza totais do orçamento)

**Importante:**
- **RECALCULA TUDO** (sobrescreve valores editados manualmente)
- **NÃO respeita** edições manuais de YGOR/JUCELI/MAYSA
- Usar apenas quando necessário forçar recálculo total

**Banco de dados:**
- Tabela: `Conjuntos do Orçamento`
- Atualiza: **TODOS** os campos fiscais e tributários

---

### 📌 **SuperAtualizaPecas** (Linha 4180)

**Responsabilidade:**
**Recalcula TODOS os impostos** de TODAS as peças do orçamento (força recálculo completo).

**Quando é chamada:**
- Usuário pressiona **F2** manualmente no grid de peças
- Botão/menu "Recalcular Impostos"

**O que faz:**

1. **Para CADA peça do orçamento:**
   - Busca dados do produto/peça
   - Recalcula **CST**
   - Recalcula **CFOP**
   - Recalcula **Base de Cálculo ICMS**
   - Recalcula **Valor ICMS**
   - Recalcula **Alíquota ICMS**
   - Recalcula **Percentual de Redução**
   - Recalcula **Valor IPI**
   - Recalcula **Alíquota IPI**
   - Recalcula **Diferido**
   - Recalcula **PIS** (com redução 48,1%)
   - Recalcula **COFINS** (com redução 48,1%)
   - Recalcula **IVA**
   - Recalcula **Base de Cálculo ST**
   - Recalcula **Valor ICMS ST**
   - Recalcula **Alíquota ICMS ST**
   - Recalcula **Valor do Tributo**

2. **Tratamento Entrega Futura:**
   - Aplica lógica específica para entrega futura
   - Aplica reduções regionais quando aplicável

3. **Atualiza valores no banco de dados**

**Chama ao final:**
- `AtualizaValoresPecas` (distribui desconto/frete)
- `AjustaValores` (atualiza totais do orçamento)

**Importante:**
- **RECALCULA TUDO** (sobrescreve valores editados manualmente)
- **NÃO respeita** edições manuais de YGOR/JUCELI/MAYSA
- Usar apenas quando necessário forçar recálculo total

**Banco de dados:**
- Tabela: `Peças do Orçamento`
- Atualiza: **TODOS** os campos fiscais e tributários

---

## ROTINAS DE REPROCESSAMENTO EM LOTE

### 📌 **RecalcularImpostosTodos** (Linha 11899)

**Responsabilidade:**
Orquestra um recálculo completo de todos os itens do orçamento ativo, reutilizando as funções individuais (`ProcessaXXX`) e os rateios financeiros.

**Como dispara:**
- Atalho `Ctrl + F11` na tela do orçamento.
- Pode ser associado a botões/menus de “Recalcular Impostos”.

**O que faz:**
1. Valida a existência de orçamento ativo e inicia transação (`vgDb.BeginTrans`).
2. Conta o total de itens (`ContarTotalItens`) para compor a barra de progresso (`pbRecalcularImpostos`).
3. Executa `ProcessarProdutosCompleto`, `ProcessarConjuntosCompleto` e `ProcessarPecasCompleto` sequencialmente.
4. Comita a transação, dispara `AjustaValores` e oculta a barra de progresso.
5. Exibe resumo com a quantidade de itens recalculados por tipo.

**Observações:**
- Em caso de erro, realiza `RollBack` e mostra a mensagem ao usuário.
- Usa os mesmos critérios de recalculo que o fluxo normal do grid (inclui bloqueios de usuários especiais).

### 📌 **ProcessarProdutosCompleto / ProcessarConjuntosCompleto / ProcessarPecasCompleto** (Linhas 12015, 12061, 12109)

**Responsabilidade:**
Simulam a inclusão/edição de cada item, iterando sobre o recordset correspondente e chamando `ProcessaProdutos`, `ProcessaConjuntos` ou `ProcessaPecas`.

**Detalhes principais:**
- Atualizam a barra de progresso para cada item processado.
- Respeitam retornos das funções `ProcessaXXX`; qualquer falha encerra o loop mantendo o contador.
- Reiniciam o recordset no início e percorrem até `EOF`, garantindo que itens inseridos fora do grid (importações) também sejam recalculados.

### 📌 **ContarTotalItens / ContarItens** (Linhas 11944 e 11966)

**Responsabilidade:**
Fornecem métricas para o recálculo em lote.

**Uso prático:**
- `ContarTotalItens` soma `RecordCount` dos recordsets globais de produtos, conjuntos e peças.
- `ContarItens` devolve o `RecordCount` para um tipo específico, usado para o resumo exibido ao final.

---

## FUNÇÃO AJUSTAVALORES

### 📌 **AjustaValores** (Linha 1247)

**Responsabilidade:**
Calcula e atualiza os **totais gerais do orçamento** (somas de impostos, valores, bases de cálculo).

**Quando é chamada:**
- Após incluir/editar/excluir itens nos grids
- Após alterar campos do cabeçalho (data, cliente, etc.)
- Múltiplos pontos do sistema

**O que faz:**

1. **Atualiza campos opcionais:**
   - `Tipo`
   - `Fechamento`

2. **~~REMOVIDO: NÃO chama mais AtualizaValores~~**
   - ~~Anteriormente chamava `AtualizaValoresProdutos`, `AtualizaValoresPecas`, `AtualizaValoresConjuntos`~~
   - ~~Isso foi REMOVIDO para evitar recálculo desnecessário ao trocar data, cliente, etc.~~

3. **Calcula totais de impostos:**
   - Soma IPI de Produtos, Conjuntos, Peças
   - Soma ICMS de Produtos, Conjuntos, Peças
   - Soma ICMS ST de Produtos, Conjuntos, Peças
   - Soma Base de Cálculo de Produtos, Conjuntos, Peças
   - Soma Base ST de Produtos, Conjuntos, Peças
   - Soma PIS de Produtos, Conjuntos, Peças
   - Soma COFINS de Produtos, Conjuntos, Peças
   - Soma Tributos de Produtos, Conjuntos, Peças

4. **Calcula valores totais:**
   - Valor Total Produtos
   - Valor Total Conjuntos
   - Valor Total Peças
   - Valor Total Serviços
   - Valor Total Produtos Usados
   - Valor Total Conjuntos Usados
   - Valor Total Peças Usadas

5. **Calcula Valor Total do Orçamento:**
   ```
   Valor Orçamento =
      Produtos + Conjuntos + Peças + Serviços +
      IPI Produtos + IPI Conjuntos + IPI Peças +
      ICMS ST Produtos + ICMS ST Conjuntos + ICMS ST Peças +
      Frete - Desconto
   ```

6. **Atualiza campos totalizadores na tabela Orçamento:**
   - `Valor do IPI`
   - `Valor do ICMS`
   - `Valor ICMS ST`
   - `Valor da Base de Cálculo`
   - `Base de Cálculo ST`
   - `Valor do PIS`
   - `Valor do Cofins`
   - `Valor do Tributo`
   - `Valor Total Produtos`
   - `Valor Total Conjuntos`
   - `Valor Total Peças`
   - `Valor Total Serviços`
   - `Valor do Orçamento`

**Importante:**
- **NÃO recalcula impostos dos itens**
- **NÃO distribui desconto/frete** (removido)
- Apenas **soma e totaliza** valores já calculados
- Executada após qualquer alteração no orçamento

**Banco de dados:**
- Tabela: `Orçamento`
- Atualiza: Apenas campos totalizadores (somas)

---

## FUNÇÃO CALCULAIMPOSTO

### 📌 **CalculaImposto** (Função externa - IRRIG.BAS ou módulo global)

**Responsabilidade:**
Função genérica que calcula **um tipo específico de imposto** baseado em parâmetros.

**Parâmetros:**
```vb
CalculaImposto(
   Sequencia_Produto,      ' ID do produto/conjunto/peça
   Sequencia_Geral,        ' ID do cadastro geral (cliente/fornecedor)
   Tipo_Imposto,           ' Código do tipo de imposto
   Tipo_Item,              ' 1=Produto, 2=Conjunto, 3=Peça
   Valor_Base,             ' Valor base para cálculo
   Valor_Adicional,        ' Valor adicional (IPI, etc.)
   Sequencia_Propriedade,  ' ID da propriedade (produtor rural)
   NCM,                    ' Classificação fiscal
   Parametro_Adicional,    ' Parâmetro extra
   UF                      ' Estado destino
)
```

**Tipos de Imposto (Tipo_Imposto):**

| Código | Imposto | Retorna |
|--------|---------|---------|
| 1 | CFOP | Código Fiscal de Operação |
| 2 | Percentual Redução | % de redução ICMS |
| 3 | Alíquota ICMS | % ICMS |
| 4 | Alíquota IPI | % IPI |
| 5 | CST | Código Situação Tributária |
| 6 | Base Cálculo ICMS | Valor da base |
| 7 | Valor ICMS | Valor do imposto |
| 8 | Valor IPI | Valor do imposto |
| 9 | Diferido | Booleano (True/False) |
| 10 | Valor PIS | Valor do imposto |
| 11 | Valor COFINS | Valor do imposto |
| 12 | IVA | Índice Valor Agregado |
| 13 | Base Cálculo ST | Valor base ST |
| 14 | Valor ICMS ST | Valor ST |
| 15 | Alíquota ICMS ST | % ICMS ST |

**Lógica Interna:**
- Consulta tabelas de **tributação** (Regras Fiscais, NCM, UF, Tipo Operação)
- Aplica **exceções fiscais** (benefícios, isenções, reduções)
- Calcula impostos conforme **legislação vigente**
- Considera **Produtor Rural** (diferimento, isenções)
- Trata **Entrega Futura** (CFOP específico, CST 90)

**Importante:**
- Função **CENTRAL** de cálculo tributário
- Chamada por todas as funções `Processa` e `SuperAtualiza`
- **NÃO** altera banco de dados diretamente
- Apenas retorna o **valor calculado**

---

## FLUXO DE EXECUÇÃO

### 🔄 **Fluxo Normal - Edição de Item no Grid (Usuário Comum)**

```
1. Usuário edita quantidade de um produto no grid
   ↓
2. Grid_AfterUpdateRecord dispara
   ↓
3. ExecutaGrid(PROCESSOS_DIRETOS) é chamado
   ↓
4. ProcessaProdutos() executa:
   - Atualiza Valor Total
   - Recalcula TODOS os impostos (CST, CFOP, ICMS, IPI, PIS, COFINS, ST, etc.)
   - Grava no banco
   ↓
5. AtualizaValoresProdutos() executa (usuário comum):
   - Distribui desconto/frete proporcionalmente
   - Recalcula PIS/COFINS com base ajustada
   - Grava no banco
   ↓
6. AjustaValores() executa:
   - Calcula totais gerais
   - Atualiza cabeçalho do orçamento
```

---

### 🔄 **Fluxo Edição Manual - YGOR/JUCELI/MAYSA**

```
1. YGOR edita manualmente "Alíquota ICMS" de 18% para 12%
   ↓
2. Grid_AfterUpdateRecord dispara
   ↓
3. ExecutaGrid(PROCESSOS_DIRETOS) é chamado
   ↓
4. ProcessaProdutos() executa:
   - Atualiza Valor Total
   - Verifica: vgPWUsuario = "YGOR"? SIM
   - ❌ PULA recálculo de impostos (mantém 12%)
   - Grava no banco
   ↓
5. AtualizaValoresProdutos() NÃO executa:
   - Verifica: vgPWUsuario = "YGOR"? SIM
   - ❌ PULA distribuição desconto/frete
   ↓
6. AjustaValores() executa:
   - Calcula totais gerais (usa valor 12% editado)
   - Atualiza cabeçalho do orçamento
```

**Resultado:** Alíquota ICMS mantém 12% editado manualmente.

---

### 🔄 **Fluxo Alteração Fora do Grid (Qualquer Usuário)**

```
1. Usuário altera DATA do orçamento
   ↓
2. Evento de alteração do campo dispara
   ↓
3. AjustaValores() é chamado
   ↓
4. ~~AtualizaValores NÃO é mais chamado~~ (REMOVIDO)
   ↓
5. Apenas calcula totais gerais:
   - Soma impostos existentes
   - Atualiza totalizadores
   ↓
6. ❌ NÃO recalcula grids
```

**Resultado:** Data alterada, grids mantêm valores inalterados.

---

### 🔄 **Fluxo F2 Manual (Qualquer Usuário)**

```
1. Usuário pressiona F2 no grid de produtos
   ↓
2. SuperAtualizaProdutos() executa
   ↓
3. Para CADA produto:
   - Recalcula TODOS os impostos (ignora edições manuais)
   - Grava no banco
   ↓
4. AtualizaValoresProdutos() executa:
   - Distribui desconto/frete
   - Recalcula PIS/COFINS
   ↓
5. AjustaValores() executa:
   - Atualiza totais gerais
```

**Resultado:** Todos os valores recalculados (SOBRESCREVE edições manuais de YGOR/JUCELI/MAYSA).

---

## REGRAS DE USUÁRIO

### 👤 **Usuários Autorizados: YGOR, JUCELI, MAYSA**

**Permissões:**
- ✅ Podem **editar manualmente** colunas fiscais (11-31) nos grids
- ✅ Valores editados **NÃO são recalculados** automaticamente
- ✅ Alterações fora do grid **NÃO recalculam** itens do grid
- ⚠️ Pressionar **F2 manual** SOBRESCREVE edições (força recálculo)

**Colunas Editáveis (Grids 0, 1, 3):**
- Coluna 11-31: Todos os campos fiscais e financeiros
  - Base de Cálculo ICMS
  - Valor ICMS
  - Valor IPI
  - Alíquota ICMS
  - Alíquota IPI
  - Diferido
  - Percentual Redução
  - IVA
  - Base ST
  - Valor ICMS ST
  - Alíquota ICMS ST
  - Bc PIS
  - Aliq PIS
  - Valor PIS
  - Bc COFINS
  - Aliq COFINS
  - Valor COFINS
  - Valor Tributo
  - Valor Desconto
  - Valor Frete

**Quando NÃO Recalcula:**
- ✅ Editar item no grid → NÃO recalcula impostos
- ✅ Alterar data orçamento → NÃO recalcula grids
- ✅ Alterar cliente → NÃO recalcula grids
- ✅ Alterar desconto/frete → NÃO recalcula grids

**Quando RECALCULA (cuidado!):**
- ⚠️ Pressionar **F2** manual → RECALCULA TUDO (sobrescreve edições)

---

### 👤 **Usuários Comuns (Outros)**

**Permissões:**
- ❌ **NÃO** podem editar colunas fiscais (11-31) - campos **bloqueados**
- ✅ Podem editar quantidade, valor unitário, descrição
- ✅ Edição dispara recálculo automático

**Quando Recalcula:**
- ✅ Editar quantidade → Recalcula impostos automaticamente
- ✅ Editar valor unitário → Recalcula impostos automaticamente
- ✅ Incluir novo item → Calcula impostos automaticamente
- ⚠️ Alterar data/cliente → **NÃO** recalcula grids (comportamento novo)

**Colunas Bloqueadas:**
- Coluna 11-31 (Grids 0, 1, 3): Campos fiscais
- Coluna 2-5 (Grid 2 - Parcelamento): Campos financeiros
- Coluna 5 (Grid 4 - Serviços): Campo fiscal

---

## ORÇAMENTO AVULSO

### 📋 **Flag: Orçamento![Orçamento Avulso]**

**Quando `Orçamento Avulso = True`:**
- ❌ **NENHUM** cálculo automático é feito
- ✅ Todos os usuários podem editar **livremente**
- ✅ Sistema **não sobrescreve** valores editados
- ✅ Útil para orçamentos **importados** ou **especiais**

**Comportamento:**
- `ProcessaProdutos` → Pula todo o bloco de cálculo
- `ProcessaPecas` → Pula todo o bloco de cálculo
- `ProcessaConjuntos` → Pula todo o bloco de cálculo
- `AtualizaValores` → NÃO é executado

**Quando usar:**
- Orçamentos com tributação especial
- Orçamentos importados de outros sistemas
- Casos onde impostos já foram calculados externamente

---

## OBSERVAÇÕES FINAIS

### ⚠️ **IMPORTANTE - Recálculo Automático**

**REMOVIDO em 10/10/2025:**
- `AjustaValores` **NÃO chama mais** `AtualizaValores`
- Alterações fora do grid **NÃO recalculam** itens

**Motivo:**
- Evitar recálculo desnecessário ao alterar data, cliente, observações
- Melhorar performance do sistema
- Evitar sobrescrever valores editados manualmente

---

### 🔧 **Manutenção e Debugging**

**Para debugar problemas de cálculo:**

1. **Verificar qual função está sendo chamada:**
   - Adicionar `Debug.Print` no início de cada função
   - Verificar se `vgPWUsuario` está correto

2. **Verificar condições de recálculo:**
   - `Orçamento![Orçamento Avulso]` = ?
   - `vgPWUsuario` = ?
   - Qual evento disparou?

3. **Verificar ordem de execução:**
   - `Processa` → `AtualizaValores` (usuário comum)
   - `Processa` → Pula `AtualizaValores` (YGOR/JUCELI/MAYSA)

4. **Verificar valores no banco:**
   - Consultar tabelas diretamente
   - Comparar antes/depois da edição

---

### 📝 **Histórico de Alterações**

**10/10/2025:**
- ✅ Adicionado controle de usuário (YGOR, JUCELI, MAYSA)
- ✅ Removido recálculo automático em `AjustaValores`
- ✅ Removido `SendK(vbKeyF2)` de `ProcessaXXX`
- ✅ Funções `AtualizaValores` só executam para usuários comuns
- ✅ Bloqueio de edição de colunas fiscais para usuários não autorizados

---

**Fim da Documentação**
*Atualizado em: 10/10/2025*
*Responsável: Assistente Claude*
