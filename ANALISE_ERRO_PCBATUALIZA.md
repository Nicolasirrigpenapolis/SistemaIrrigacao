# Análise do Erro "Erro ao recarregar orçamento" - PcbAtualizaOrc

## Data da Análise
05 de Novembro de 2025

## Problema Relatado
Ao clicar no `pcbatualiza` está aparecendo o erro "Erro ao recarregar orçamento". O sistema não consegue recarregar o orçamento.

## Tabelas e Colunas Acessadas

### 1. Procedimento `PcbAtualizaOrc_Click()`

#### Tabelas DELETE:
- **[Itens Pendentes]**
  - Coluna: `[Seqüência do Orçamento]`
  
- **[Receita Primaria]**
  - Coluna: `[Seqüência do Orçamento]`

### 2. Procedimento `MySQL()` - Chamado por PcbAtualizaOrc_Click

#### Tabela Access (vgDb):
- **[Itens pendentes]**
  - Colunas SELECT: `*` (todas)
  - Colunas WHERE: `[Seqüência Do Orçamento]`
  - Colunas INSERT: 
    - `[Seqüência do Orçamento]`
    - `[Sequencia do Item]`
    - `[Seqüência do Conjunto]`
    - `[Seqüência do Produto]`
    - `Quantidade`
    - `[Valor Total]`
    - `[Valor Unitário]`
    - `Tp`

#### Tabelas SQL Server (cnGas):
- **[Peças do Orçamento]**
  - Colunas: `[Seqüência do Produto]`, `Quantidade`, `[Valor Total]`, `[Valor Unitário]`, `[Seqüência do Orçamento]`
  
- **Orçamento**
  - Colunas: `[Seqüência do Orçamento]`
  
- **Produtos**
  - Colunas: `[Seqüência Do Produto]`, `Descrição`
  
- **[Conjuntos do Orçamento]**
  - Colunas: `[Seqüência Do Conjunto]`, `Quantidade`, `[Valor Total]`, `[Valor Unitário]`, `[Seqüência do Orçamento]`
  
- **Conjuntos**
  - Colunas: `[Seqüência Do Conjunto]`, `Descrição`
  
- **[Produtos do Orçamento]**
  - Colunas: `[Seqüência Do Produto]`, `Quantidade`, `[Valor Total]`, `[Valor Do IPI]`, `[Valor Unitário]`, `[Seqüência do Orçamento]`

### 3. Procedimento `ReceitaP()` - Chamado por PcbAtualizaOrc_Click

#### Tabelas Acessadas:
- **[Receita Primaria]**
  - Colunas SELECT: `*` (todas)
  - Colunas INSERT/UPDATE:
    - `[Seqüência do Produto]`
    - `[Seqüência da Matéria Prima]`
    - `[Seqüência do Orçamento]`
    - `Quantidade`
    - `[Sequencia do Item]`
    - `[Id do Pedido]`
    - `Pagto`
    - `[Qtde Recebida]`
    - `[Qtde Restante]`
    - `[Qtde Total]`
    - `Localização`
    - `Situação`
    - `[Seqüência do Conjunto]`
    - `[Sequencia Produto Principal]`

- **Produtos**
  - Colunas: `[Seqüência do Produto]`, `Localização`, `[Material Adquirido de Terceiro]`, `[Nao sair no checklist]`, `[Mostrar Receita Secundaria]`, `Descrição`, `[Tipo do Produto]`, `[Receita Conferida]`

- **[Itens Pendentes]**
  - Colunas: `[Seqüência do Produto]`, `[Seqüência do Orçamento]`, `Quantidade`, `Lance`, `situação`, `[Seqüência do Conjunto]`, `TP`

- **[Matéria Prima]**
  - Colunas: `[Seqüência do Produto]`, `[Seqüência da Matéria Prima]`, `[Quantidade de Matéria Prima]`

- **[Itens do Conjunto]**
  - Colunas: `[Seqüência do Produto]`, `[Seqüência do Conjunto]`, `[Material Adquirido de Terceiro]`, `[Mostrar Receita Secundaria]`, `[Quantidade Do Produto]`, `Descrição`

- **Conjuntos**
  - Colunas: `[Seqüência do Conjunto]`, `[Receita Conferida]`

## Possíveis Causas do Erro

### 1. **Problema de Conexão SQL Server** (MAIS PROVÁVEL)
O procedimento `MySQL()` tenta conectar ao SQL Server usando diferentes strings de conexão:
- **Em IDE no computador DESKTOP-CTAJU78**: `DESKTOP-CTAJU78\SQLEXPRESS02`
- **Em IDE outros computadores**: `DESKTOP-CHS14C0\SQLIRRIGACAO`
- **Em produção**: `SRVSQL\SQLEXPRESS` com usuário `ygor` e senha `5139249_`

**Verificações necessárias:**
- O servidor SQL está acessível?
- As credenciais estão corretas?
- O banco `IRRIGACAO` existe no SQL Server?

### 2. **Colunas Inexistentes ou com Nomes Diferentes**
Possíveis problemas:
- Coluna `[Seqüência do Orçamento]` vs `[Seqüência Do Orçamento]` (maiúsculas/minúsculas)
- Coluna `[Sequencia do Item]` sem acento vs com acento
- Coluna `[Id do Pedido]` vs `[Id Do Pedido]`

### 3. **Tabelas Ausentes**
Verificar se existem no SQL Server:
- `[Peças do Orçamento]`
- `Orçamento`
- `[Conjuntos do Orçamento]`
- `[Produtos do Orçamento]`

### 4. **Problema nas Funções Auxiliares**
O código chama funções que podem não estar funcionando:
- `MostraCompra()`
- `MostraFinanceiro()`
- `QtdeRecebida()`
- `QtdeRestante()`
- `QtdeTotal()`

## Melhorias Implementadas

### 1. Tratamento de Erro Aprimorado em `PcbAtualizaOrc_Click()`
Adicionada variável `etapaErro` para identificar em qual etapa o erro ocorre:
- iniciando transação
- removendo Itens Pendentes
- removendo Receita Primária
- executando MySQL (inserindo pendentes)
- executando ReceitaP (inserindo receita primária)
- confirmando transação
- recarregando contas
- atualizando form de Receita Primária

A mensagem de erro agora mostra:
- A etapa onde ocorreu o erro
- A descrição do erro
- O código do erro

### 2. Tratamento de Erro em `MySQL()`
Adicionado:
- `On Error GoTo ErroMySQL`
- Fechamento seguro de conexões em caso de erro
- Propagação do erro com fonte identificada

### 3. Tratamento de Erro em `ReceitaP()`
Adicionado:
- `On Error GoTo ErroReceitaP`
- Propagação do erro com fonte identificada
- Reset do mouse pointer em caso de erro

## Próximos Passos para Diagnóstico

1. **Execute o sistema e tente clicar em `pcbatualiza`**
2. **Anote a mensagem de erro completa**, especialmente:
   - Qual etapa estava executando
   - O código do erro
   - A descrição do erro
3. **Verifique:**
   - Se o SQL Server está acessível
   - Se as tabelas existem no SQL Server
   - Se as colunas têm os nomes corretos

## Exemplo de Como Testar

```vb
' Após clicar em pcbatualiza, você verá uma mensagem como:
' "Erro ao recarregar orçamento na etapa: executando MySQL (inserindo pendentes)
'
' Erro: [Descrição do erro] (Código: [Número])"
```

Com essa informação, será possível identificar exatamente qual tabela ou coluna está causando o problema.
