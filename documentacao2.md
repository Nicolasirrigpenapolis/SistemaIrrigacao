# Documentação: Análise do Valor Unitário no Orçamento (12/11/2025)

Este documento analisa como o "Valor Unitário" funciona hoje no sistema, documenta os problemas que estávamos tendo e propõe solução com plano de implementação.

---

## Visão geral

**Arquivo principal**: `IRRIG/ORCAMENT.FRM`

**Grids envolvidos**:
- Grid(0): Conjuntos - coluna 7 (Valor Unitário)
- Grid(1): Peças - coluna 9 (Valor Unitário)
- Grid(3): Produtos - coluna 10 (Valor Unitário)

**Funções auxiliares principais**:
- **Info\***: `InfoConjuntos`, `InfoPecas`, `InfoProdutos` - buscam dados do cadastro (referência)
- **Valida\***: `ValidaConjx`, `ValidaPecasx`, `ValidaProdx` - validam valor mínimo
- **Processa\***: `ProcessaConjunto`, `ProcessaPeca`, `ProcessaProduto` - gravam e calculam impostos

---

## Como funciona hoje

### Grid(0) — Conjuntos (coluna 7)

**Exibição (CONTEUDODACOLUNA)**:
- Coluna 7 exibe o valor via `InfoConjuntos("Valor Unitário")` → busca `ConjuntoAux![Valor Total]` do cadastro

**Edição/leitura (IniApDaCol)**:
- Carrega o valor de `ColumnValue(7)` → valor armazenado no orçamento

**Validações**:
- `ValidaConjx(...)` exige `Valor_Unitario > 0` e `>= ConjuntoAux![Valor Total]`
- Se não atende, chama `PermissaoConj(...)` para exceção (usuários: YGOR, WAGNER, JERONIMO, ALEXANDRE, CESAR, JUCELI, MAYSA)

**Processamento**:
- `ProcessaConjunto(...)` grava e calcula tributos

---

### Grid(1) — Peças (coluna 9)

**Exibição (CONTEUDODACOLUNA)**:
- Coluna 9 exibe o valor via `InfoPecas("Valor Unitário")` → busca `PecaAux![Valor Total]` do cadastro

**Edição/leitura (IniApDaCol)**:
- Carrega o valor de `ColumnValue(9)` → valor armazenado no orçamento

**Validações**:
- `ValidaPecasx(...)` exige `Valor_Unitario > 0` e `>= ProdutoAux![Valor Total]`
- Se não atende, chama `PermissaoPecas(...)` para exceção (mesma lista de usuários)

**Processamento**:
- `ProcessaPeca(...)` grava e recalcula impostos

---

### Grid(3) — Produtos (coluna 10)

**Exibição (CONTEUDODACOLUNA)**:
- Coluna 10 exibe o valor via `InfoProdutos("Valor Unitário")` → busca do cadastro
- **Regra especial**: Se produto do grupo 20 E existe MP 43602 → retorna `Valor de Custo * 3.5`
- Caso contrário → retorna `ProdutoAux![Valor Total]`

**Edição/leitura (IniApDaCol)**:
- Carrega o valor de `ColumnValue(10)` → valor armazenado no orçamento

**Validações**:
- `ValidaProdx(...)` exige `Valor_Unitario > 0` e `>= ProdutoAux![Valor Total]`
- ⚠️ **Importante**: validação sempre usa `Valor Total`, IGNORA a regra do custo*3.5
- Se não atende, chama `Permissao(...)` para exceção

**Processamento**:
- `ProcessaProduto(...)` grava e calcula impostos
- Se `Edicao_Manual_Impostos = True`, segue caminho manual

---

## Problemas identificados

### 1. Divergência entre exibição e cálculo

**O problema**:
- A coluna "Valor Unitário" EXIBE sempre o valor do cadastro (via `Info*`)
- Mas os CÁLCULOS usam o valor do orçamento (via `ColumnValue`)
- Usuário vê um número na tela, mas o sistema usa outro nos cálculos

**Consequência**:
- Usuário edita para R$ 95,00
- Tela continua mostrando R$ 100,00 (do cadastro)
- Mas total calcula com R$ 95,00 (do orçamento)
- Confusão: "não salvou minha edição!"

---

### 2. Troca de item mantém preço errado

**O problema**:
- Não existe código para resetar `ColumnValue` quando troca de item
- Sistema mantém o preço do item anterior em `ColumnValue`
- Mas exibe o preço do novo item via `Info*`

**Consequência**:
- Seleciona Produto A (R$ 100), edita para R$ 95
- Troca para Produto B (R$ 200)
- Sistema usa R$ 95 no cálculo (preço errado!)
- Mas exibe R$ 200 na tela
- Total fica errado, usuário não entende

**⚠️ Cuidados ao resolver**:
- Precisa permitir edição manual do valor unitário
- Precisa bloquear quando digitar valor menor que o cadastrado (ou pedir permissão)
- Reset só deve acontecer na TROCA de item, não durante edição normal

---

### 3. Produtos grupo 20 - divergência de regras

**O problema**:
- `InfoProdutos` retorna `Custo * 3.5` quando produto é grupo 20 E tem MP 43602
- Mas `ValidaProdx` sempre compara com `Valor Total` (ignora a regra custo*3.5)

**Consequência**:
- Tela mostra mínimo de R$ 175 (custo*3.5)
- Mas validação aceita R$ 160 (acima de Valor Total R$ 150)
- Inconsistência: o "mínimo exibido" não é o "mínimo validado"

---

### 4. Falta de clareza visual

**O problema**:
- Não existe indicação visual do "valor de referência" vs "valor orçado"
- Validação diz "não pode ser menor que o valor do sistema", mas não mostra qual é esse valor

---

### 5. Falta de auditoria

**O problema**:
- Quando usuário autorizado vende abaixo do mínimo, não há registro do motivo
- Lista de usuários hardcoded no código
- Impossível rastrear vendas abaixo do mínimo depois

---

## Proposta de solução

### Separar visualmente "Ref. Sistema" e "Valor Unitário"

**Criar duas colunas distintas**:

1. **"Ref. Sistema"** (nova, somente leitura):
   - Exibe o valor de referência do cadastro (mínimo permitido)
   - Vem de `Info*("Valor Unitário")`
   - Usuário NÃO pode editar

2. **"Valor Unitário"** (atual, editável):
   - Exibe o valor DO ORÇAMENTO (o que está sendo usado no cálculo)
   - Vem de `ColumnValue`
   - Usuário PODE editar

**Regras**:
- Na inclusão: "Valor Unitário" = "Ref. Sistema" automaticamente
- Ao trocar item: "Valor Unitário" reseta para "Ref. Sistema" do NOVO item
- Na edição manual: Usuário pode digitar qualquer valor
- Validação: "Valor Unitário" >= "Ref. Sistema" → se menor, bloqueia OU pede permissão

**Benefícios**:
- ✅ Usuário vê exatamente o que está sendo calculado
- ✅ Clareza: duas colunas mostram referência vs orçado
- ✅ Preço não fica "preso" ao trocar item
- ✅ Continua permitindo edição manual
- ✅ Validação clara com mensagem informando ambos os valores

---

## Ordem de implementação

1. **Fase 1**: Grid 3 - Produtos (mais complexo, tem regra grupo 20)
2. **Fase 2**: Grid 1 - Peças (replicar)
3. **Fase 3**: Grid 0 - Conjuntos (replicar)
4. **Fase 4** (opcional): Auditoria com log de vendas abaixo do mínimo

**Importante**: Testar completamente cada fase antes de prosseguir para a próxima.
