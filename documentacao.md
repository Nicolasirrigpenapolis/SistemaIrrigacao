# Documentação: Análise do Valor Unitário no Orçamento (12/11/2025)


Es**O problema**O problema observado**:
- Nas colunas de "Valor Unitário", a exibição usa SEMPRE o valor de referência do cadastro (via `Info*`)
- Já o cálculo do total e as validações usam a variável `Valor_Unitario` carregada de `ColumnValue` (valor que está na linha do orçamento)

**Consequência prática**:
- Usuário edita o "Valor Unitário" para um valor negociado (ex: R$ 95,00)
- A UI continua exibindo o valor de referência do cadastro (ex: R$ 100,00)
- Mas o total mostrado é calculado a partir do valor editado (R$ 95,00)
- **Confusão**: "não salvou minha edição!"

### 2. Troca de item mantém preço errado

**O problema observado**:
- Não existe código para resetar `ColumnValue` quando troca de item
- Sistema mantém o preço do item anterior em `ColumnValue`
- Mas exibe o preço do novo item via `Info*`

**Consequência prática**:
- Seleciona Produto A (R$ 100), edita para R$ 95
- Troca para Produto B (R$ 200)
- Sistema usa R$ 95 no cálculo (preço do Produto A!)
- Mas exibe R$ 200 na tela (preço do Produto B)
- Total fica errado, validação falha, usuário não entende

### 3. Produtos grupo 20 - divergência de regras

**O problema observado**:vado**:
- Nas colunas de "Valor Unitário", a exibição usa SEMPRE o valor de referência do cadastro (via `Info*`)
- Já o cálculo do total (ex.: `Quantidade * Valor_Unitario`) e as validações usam a variável `Valor_Unitario` carregada de `ColumnValue` (ou seja, o valor que realmente está na linha do orçamento)

**Consequência prática**:
- Usuário edita o "Valor Unitário" para um valor negociado (ex: R$ 95,00)
- A UI continua exibindo o valor de referência do cadastro (ex: R$ 100,00)
- Mas o total mostrado é calculado a partir do valor editado de `ColumnValue` (R$ 95,00)
- **Confusão**: "não salvou minha edição!"

### 2. Troca de item mantém preço errado

**O problema observado**:umento analisa como o "Valor Unitário" funciona hoje no sistema, documenta os problemas que estávamos tendo e propõe solução com plano de implementação.

## Visão geral

Arquivo principal: `IRRIG/ORCAMENT.FRM`

Grids envolvidos no orçamento:
- Grid(0): Conjuntos
- Grid(1): Peças
- Grid(3): Produtos

Funções auxiliares chave:
- Busca de dados (cadastro / referência): `InfoConjuntos`, `InfoPecas`, `InfoProdutos`
- Validações de valor mínimo (não pode ser menor que o do sistema): `ValidaConjx`, `ValidaPecasx`, `ValidaProdx`
- Processamento/gravação e impostos: `ProcessaConjunto`, `ProcessaPeca`, `ProcessaProduto`

Observação importante: Hoje as colunas de “Valor Unitário” em cada grid exibem o valor de referência do cadastro (via `Info*`) durante a renderização da célula, enquanto as validações e totais usam o valor do orçamento carregado de `ColumnValue`. Isso gera inconsistências visuais e de uso (detalhes abaixo).

---

## Como funciona hoje, por grid

### Grid(0) — Conjuntos
- Exibição (CONTEÚDO DA COLUNA):
  - Coluna 7 (Valor Unitário) exibe o valor do cadastro via `InfoConjuntos("Valor Unitário")`.
  - Outras colunas derivadas também usam `InfoConjuntos` (ex.: “Sigla”, “Estoque”).
- Edição/leitura do registro:
  - `IniApDaCol` carrega o “Valor Unitário” a partir de `ColumnValue(7)` (valor armazenado na linha do orçamento).
- Validações:
  - Atividade do conjunto: `ValidaConjunto(...)` bloqueia conjunto inativo (“Conjunto INATIVO!”).
  - Quantidade: exige `Quantidade > 0` (mensagem “Quantidade inválido!”).
  - `ValidaConjx(...)` exige `Valor_Unitario > 0` e `>= ConjuntoAux![Valor Total]` (valor de referência do cadastro). Caso não atenda, pode acionar `PermissaoConj(...)` para exceção com permissão.
  - Mensagem quando inválido: “Valor Unitário inválido!” (coluna 7).
- Gravação e impostos:
  - `ProcessaConjunto(...)` atualiza os campos e calcula tributos conforme parâmetros (sem caminho de edição manual específico neste grid).

### Grid(1) — Peças
- Exibição (CONTEÚDO DA COLUNA):
  - Coluna 9 (Valor Unitário) exibe o valor do cadastro via `InfoPecas("Valor Unitário")`.
- Edição/leitura do registro:
  - `IniApDaCol` usa `ColumnValue(9)` para “Valor Unitário” (valor do orçamento).
- Validações:
  - Atividade da peça: `ValidaProduto3(...)` bloqueia peça inativa (mensagem “Impossivel Peça Inativa!”).
  - Cadastro do item: `PodeVenderPecas(...)` bloqueia item com cadastro incompleto (mensagem “Impossivel Orçar Cadastro do Item Incompleto!”).
  - Quantidade: exige `Quantidade > 0` (mensagem “Quantidade inválido!”).
  - `ValidaPecasx(...)` exige `Valor_Unitario > 0` e `>= ProdutoAux![Valor Total]` (referência do cadastro da peça). Caso não atenda, pode acionar `PermissaoPecas(...)`.
  - Mensagem quando inválido: “Valor Unitário não pode ser menor que o Valor do Sistema!(Valor Unitário Invalido)” (coluna 9).
- Gravação e impostos:
  - `ProcessaPeca(...)` atualiza os campos da linha e recalcula impostos associados.

### Grid(3) — Produtos
- Exibição (CONTEÚDO DA COLUNA):
  - Coluna 10 (Valor Unitário) exibe o valor do cadastro via `InfoProdutos("Valor Unitário")`.
  - Regra especial em `InfoProdutos("Valor Unitário")`: se o produto é do grupo 20 e existir uma matéria-prima específica (seq. 43602), o valor de referência vira `Valor de Custo * 3,5`; caso contrário usa `ProdutoAux![Valor Total]`.
- Edição/leitura do registro:
  - `IniApDaCol` usa `ColumnValue(10)` (valor do orçamento).
- Validações:
  - Atividade do produto: `ValidaProduto2(...)` bloqueia produto inativo (mensagem “Impossivel Produto Inativo!”).
  - Cadastro do item: `PodeVenderProd(...)` bloqueia item com cadastro incompleto (mensagem “Impossivel Orçar Cadastro do Item Inclompleto!”).
  - NCM: `ValidaNCM(...)` exige NCM válido (mensagem “Pedir para Contabilidade Conferir o (NCM)”).
  - Quantidade: exige `Quantidade > 0` (mensagem “Quantidade inválido!”).
  - Preço mínimo: `ValidaProdx(...)` exige `Valor_Unitario > 0` e `>= ProdutoAux![Valor Total]`. Se não atende, tenta `Permissao(...)` para exceção.
  - Mensagem quando inválido: “Valor Unitário não pode ser menor que o Valor do Sistema!(Valor Unitário Invalido)” (coluna 10).
- Gravação e impostos:
  - `ProcessaProduto(...)` grava e calcula impostos. Se `Edicao_Manual_Impostos = True`, segue um caminho “manual” atualizando Valor Unitário/Total diretamente e chamando `RecalculaProdutos`.

---

## Problemas identificados (que estávamos tendo)

### 1. Divergência entre exibição e cálculo

- Nas colunas de “Valor Unitário”, a exibição usa SEMPRE o valor de referência do cadastro (via `Info*`).
- Já o cálculo do total (ex.: `Quantidade * Valor_Unitario`) e as validações usam a variável `Valor_Unitario` carregada de `ColumnValue` (ou seja, o valor que realmente está na linha do orçamento, incluso se o usuário alterou).
- Isso significa que a interface pode mostrar um valor e calcular/validar com outro, confundindo o usuário:
  - Ex.: Usuário edita o “Valor Unitário” para um valor negociado; a UI continua exibindo o valor de referência do cadastro, mas o total mostrado é calculado a partir do valor editado (de `ColumnValue`).

Além disso, há uma divergência específica em Produtos (grupo 20):
- A coluna exibe o “Ref. Sistema” calculado como `Valor de Custo * 3,5` (via `InfoProdutos("Valor Unitário")`) quando há a matéria-prima 43602.
- Porém a validação de mínimo (`ValidaProdx`) compara contra `ProdutoAux![Valor Total]` (não contra o custo*3,5).
- Risco: o “mínimo exibido” pode não ser o mesmo “mínimo validado”, gerando aceitações ou bloqueios incoerentes. É necessário alinhar a definição única de “Valor de Referência” usada para exibir e para validar.

---

## Dificuldades atuais (dores mapeadas)

1. Confusão visual e funcional: a coluna “Valor Unitário” mostra o valor do cadastro, não o valor do orçamento. O usuário sente que “não salvou” ou que “sobrescreve”.
2. Inconsistência: total e validações usam um valor (orçado), a célula exibe outro (referência).
3. Troca de item: sem uma regra clara de reset do “Valor Unitário” do orçamento, pode ficar um valor associado a outro item; a exibição por `Info*` mascara o problema mostrando o de referência do novo item.
4. Regras de mínimo: a mensagem e validação existem, mas sem clareza visual de qual é o “mínimo do sistema” versus o “valor orçado”.
5. Caminhos especiais: em Produtos, `Edicao_Manual_Impostos` permite atualizar diretamente, ampliando a diferença de comportamento entre grids.
6. Manutenção difícil: a lógica de exibição (Info*) e a de cálculo (ColumnValue) estão desacopladas e geram efeitos colaterais.

---

## O que precisamos (requisitos)

**Resumo executivo — O problema e a solução**

**Problema central**: 
A coluna "Valor Unitário" mostra o valor de referência do cadastro (via `Info*`), mas os cálculos e validações usam o valor do orçamento (via `ColumnValue`). O usuário vê um número, mas o sistema trabalha com outro — isso gera confusão e bugs silenciosos (ex.: troca de item mantém preço do item anterior, mas exibe o preço do novo).

**Solução**:
1. Criar coluna "Ref. Sistema" (somente leitura) mostrando `Info*("Valor Unitário")`.
2. Mudar coluna "Valor Unitário" para exibir o valor real do orçamento (`ColumnValue`).
3. Preencher "Valor Unitário" automaticamente com "Ref. Sistema" na inclusão.
4. Resetar "Valor Unitário" ao trocar de item.
5. Validar que "Valor Unitário" >= "Ref. Sistema" (com permissão de exceção quando necessário).

**Benefícios**:
- Clareza total: usuário vê exatamente o que está sendo usado no cálculo.
- Segurança: preço não fica "preso" de um item anterior ao trocar.
- Auditoria: permite rastrear quando/por que vendeu abaixo do mínimo.
- Manutenção: código mais limpo, uma única fonte de verdade para cada valor.

---

## Decisões técnicas críticas a tomar ANTES de implementar

### 1. Produtos grupo 20 — Qual é o "Valor de Referência"?

**Situação atual**:
- `InfoProdutos("Valor Unitário")` retorna `Valor de Custo * 3.5` se grupo 20 + MP 43602, senão retorna `ProdutoAux![Valor Total]`.
- `ValidaProdx` sempre compara com `ProdutoAux![Valor Total]` (ignora a regra custo*3.5).

**Opções**:
- **A)** Manter a exibição como está, mas fazer `ValidaProdx` também usar custo*3.5 quando aplicável → garante coerência.
- **B)** Remover a regra especial custo*3.5 e sempre usar `ProdutoAux![Valor Total]` → simplifica, mas pode não refletir a política comercial desejada.

**Recomendação**: **Opção A** — manter a regra custo*3.5 (que provavelmente existe por motivo de negócio), mas aplicar também na validação. Assim "Ref. Sistema" e "mínimo permitido" serão sempre o mesmo valor.

### 2. Reset ao trocar item — Sempre ou Condicional?

**Opções**:
- **A)** Sempre resetar "Valor Unitário" = "Ref. Sistema" ao detectar troca de item.
- **B)** Só resetar se "Valor Unitário" estiver zerado/vazio.
- **C)** Perguntar ao usuário se quer manter o preço anterior ou usar o novo (popup).

**Recomendação**: **Opção A** — sempre resetar. É mais seguro e evita carregar preço errado. Se o usuário quer um preço customizado, ele edita após selecionar o item correto.

### 3. Permissão abaixo do mínimo — Como funciona?

**Situação atual** (verificada no código):
- `Permissao` (Produtos): permite se não houver matéria-prima cadastrada OU se usuário for YGOR, WAGNER, JERONIMO, ALEXANDRE, CESAR, JUCELI ou MAYSA.
- `PermissaoPecas` (Peças): permite se não houver matéria-prima cadastrada OU se usuário for YGOR, WAGNER, JERONIMO, ALEXANDRE, CESAR, JUCELI ou MAYSA.
- `PermissaoConj` (Conjuntos): permite apenas se usuário for YGOR, WAGNER, JERONIMO, ALEXANDRE, CESAR, JUCELI ou MAYSA.

**Problemas identificados**:
- Não há justificativa/motivo registrado quando a exceção é usada.
- Não há log/auditoria das vendas abaixo do mínimo.
- Lista de usuários hardcoded no código (dificulta manutenção).
- Se produto não tem matéria-prima, permite vender a qualquer preço (pode ser intencional para produtos adquiridos de terceiros, mas não está documentado).

**Melhorias recomendadas**:
- Criar popup solicitando "Motivo da redução" quando usuário autorizado vende abaixo do mínimo.
- Registrar em tabela de auditoria: data/hora, usuário, orçamento, item, valor mínimo, valor vendido, motivo.
- Mover lista de usuários autorizados para tabela de configuração (ex.: "Usuários com Permissão Especial").
- Documentar claramente a regra de produtos sem matéria-prima.

### 4. Adicionar coluna no grid — Impacto em layout

**Consideração**:
- Adicionar coluna "Ref. Sistema" vai aumentar largura do grid.
- Alternativa: mostrar "Ref. Sistema" em tooltip/hint ao lado da coluna "Valor Unitário", sem ocupar espaço extra.

**Recomendação**: Adicionar coluna física é mais claro. Se necessário, ajustar larguras de outras colunas ou aumentar grid. Priorizar clareza sobre economia de espaço.

---

## Plano de implementação passo a passo

### FASE 0: Preparação e decisões (você precisa decidir antes de começar)

**Decisões obrigatórias**:
1. ✅ **Produtos grupo 20**: Opção A (manter custo*3.5 e alinhar validação) ou Opção B (remover regra especial)?
2. ✅ **Reset ao trocar item**: Opção A (sempre resetar) ou Opção B (só se vazio)?
3. ✅ **Layout**: Adicionar coluna física "Ref. Sistema" ou usar tooltip/hint?
4. ✅ **Auditoria**: Implementar log de vendas abaixo do mínimo agora ou depois?

---

### FASE 1: Grid 3 - Produtos (piloto)

**Passo 1.1 — Alinhar regra de "Valor de Referência" em Produtos grupo 20**

Se decisão = Opção A (manter custo*3.5):
```vb
' Em ValidaProdx (linha ~12842), substituir:
ValidaProdx = Valor_Unitario > 0 And Valor_Unitario >= ProdutoAux![Valor Total]

' Por (usar a mesma lógica de InfoProdutos):
Dim ValorRef As Double
Dim MP As New GRecordSet
Set MP = vgDb.OpenRecordSet("SELECT * From [Matéria Prima] WHERE [Seqüência do Produto] = " & Sequencia_do_Produto & " And [Seqüência da Matéria Prima] = 43602")
If ProdutoAux![Seqüência do Grupo Produto] = 20 And MP.RecordCount > 0 Then
   ValorRef = ProdutoAux![Valor de Custo] * 3.5
Else
   ValorRef = ProdutoAux![Valor Total]
End If
Set MP = Nothing
ValidaProdx = Valor_Unitario > 0 And Valor_Unitario >= ValorRef
```

**Passo 1.2 — Adicionar coluna "Ref. Sistema" no Grid 3**

1. No designer do formulário `ORCAMENT.FRM`, abrir Grid(3).
2. Adicionar nova coluna antes da coluna 10 (Valor Unitário):
   - Nome: "Ref. Sistema"
   - Somente leitura: True
   - Largura: ~80-100 pixels
3. Ajustar índices das colunas subsequentes (+1).

**Passo 1.3 — Alterar CONTEUDODACOLUNA para Grid 3**

Localizar `ExecutaGrid3`, seção `If vgOq = CONTEUDODACOLUNA Then`, substituir:
```vb
' Linha ~21443 - ANTES (Case 10):
Case 10
   vgRetVal = (InfoProdutos(..., "Valor Unitário"))
   vgColumn = -1

' DEPOIS - dividir em dois cases:
Case 10  ' Nova coluna "Ref. Sistema"
   vgRetVal = (InfoProdutos(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal, "Valor Unitário"))
   vgColumn = -1
Case 11  ' Coluna "Valor Unitário" (agora mostra o valor do orçamento)
   vgRetVal = Valor_Unitario  ' De ColumnValue, não mais Info*
   vgColumn = -1
```

**Passo 1.4 — Atualizar IniApDaCol para Grid 3**

Ajustar índices de `ColumnValue`:
```vb
' IniApDaCol em ExecutaGrid3 (linha ~21595)
' Valor_Unitario era ColumnValue(10), agora será ColumnValue(11):
Valor_Unitario = ColumnValue(11)
Valor_do_Desconto = ColumnValue(12)  ' era 11
Valor_do_Frete = ColumnValue(13)     ' era 12
' ... e assim por diante, +1 em todos
```

**Passo 1.5 — Adicionar default em DEFAULTDASCOLUNAS**

```vb
' ExecutaGrid3, seção DEFAULTDASCOLUNAS
Case 11  ' Valor Unitário
   Valor_Unitario = InfoProdutos(Bc_pis, Valor_do_Tributo, Sequencia_do_Orcamento, Sequencia_do_Produto_Orcamento, Sequencia_do_Produto, Quantidade, _
   Valor_Unitario, Valor_Total, Valor_do_IPI, Valor_do_ICMS, Aliquota_do_IPI, Aliquota_do_ICMS, _
   Percentual_da_Reducao, Diferido, Valor_da_Base_de_Calculo, Valor_do_PIS, Valor_do_Cofins, IVA, _
   Base_de_Calculo_ST, Valor_ICMS_ST, CFOP, CST, Aliquota_do_ICMS_ST, Valor_do_Frete, _
   Valor_do_Desconto, Valor_Anterior, Bc_cofins, Aliq_do_pis, Aliq_do_cofins, Peso, PesoTotal, "Valor Unitário")
   vgRetVal = Valor_Unitario
```

**Passo 1.6 — Detectar e resetar na troca de item (se decisão = sempre resetar)**

```vb
' No IniApDaCol de ExecutaGrid3, ANTES de chamar InfoProdutos:
Static SequenciaProdutoAnterior As Long

If Grid(3).Status <> ACAO_INCLUINDO Then
   If Produtos_do_Orcamento.EOF = False And Produtos_do_Orcamento.BOF = False And Produtos_do_Orcamento.RecordCount > 0 Then
      Sequencia_do_Produto_Orcamento = Produtos_do_Orcamento![Seqüência do Produto Orçamento]
      Valor_Total = Produtos_do_Orcamento![Valor Total]
      Valor_Anterior = Produtos_do_Orcamento![Valor Anterior]
      
      ' Detectar troca de item
      If SequenciaProdutoAnterior > 0 And SequenciaProdutoAnterior <> Sequencia_do_Produto Then
         ' Resetar valor unitário para o do novo item
         Valor_Unitario = InfoProdutos(..., "Valor Unitário")
      End If
      SequenciaProdutoAnterior = Sequencia_do_Produto
   End If
End If
```

**Passo 1.7 — Melhorar mensagem de validação**

```vb
' Em ExecutaGrid3, seção VALIDACOES, Case vgCol = -1 para ValidaProdx:
If vgIsValid And vgCol = -1 Then
   Dim ValorRefProduto As Double
   ' Calcular valor de referência (mesma lógica de InfoProdutos)
   ValorRefProduto = InfoProdutos(..., "Valor Unitário")
   
   vgIsValid = (ValidaProdx(...))
   If Not vgIsValid Then vgColumn = 11  ' Ajustar índice da coluna
   vgErrorMessage$ = "Valor Unitário (R$ " & Format(Valor_Unitario, "0.00") & ") não pode ser menor que o Valor de Referência do Sistema (R$ " & Format(ValorRefProduto, "0.00") & ")!"
End If
```

**Passo 1.8 — Testar Grid 3**

1. ✅ Incluir produto novo → verifica se "Ref. Sistema" e "Valor Unitário" aparecem iguais.
2. ✅ Editar "Valor Unitário" para cima → deve aceitar e recalcular impostos.
3. ✅ Editar "Valor Unitário" para baixo (abaixo do mínimo) → deve bloquear ou pedir permissão.
4. ✅ Trocar produto → deve resetar "Valor Unitário" para o do novo produto.
5. ✅ Produto grupo 20 → verificar que exibição e validação usam a mesma regra (custo*3.5).
6. ✅ Modo manual (`Edicao_Manual_Impostos = True`) → deve funcionar normalmente.
7. ✅ Entrega futura → CFOP/CST/impostos calculados corretamente.

---

### FASE 2: Grid 1 - Peças (replicar)

Seguir os mesmos passos da Fase 1, ajustando:
- Coluna: 9 → adicionar "Ref. Sistema" antes
- Função Info: `InfoPecas`
- Função Valida: `ValidaPecasx`
- Não há regra especial de grupo (mais simples que Produtos)

---

### FASE 3: Grid 0 - Conjuntos (replicar)

Seguir os mesmos passos da Fase 1, ajustando:
- Coluna: 7 → adicionar "Ref. Sistema" antes
- Função Info: `InfoConjuntos`
- Função Valida: `ValidaConjx`
- Não há regra especial de grupo (mais simples que Produtos)

---

### FASE 4: Auditoria (opcional, mas recomendado)

**Criar tabela de log**:
```sql
CREATE TABLE [Vendas Abaixo Minimo] (
   [Sequência Log] AUTOINCREMENT PRIMARY KEY,
   [Data Hora] DATETIME,
   [Usuario] VARCHAR(50),
   [Seqüência do Orçamento] LONG,
   [Tipo Item] VARCHAR(20),  -- 'Produto', 'Peça', 'Conjunto'
   [Seqüência do Item] LONG,
   [Descrição Item] VARCHAR(255),
   [Valor Referência] CURRENCY,
   [Valor Vendido] CURRENCY,
   [Motivo] MEMO
)
```

**Alterar funções de permissão**:
```vb
Private Function PermissaoComLog(ValorRef As Double, ValorVendido As Double, ...) As Boolean
   If (usuário autorizado) Then
      Dim motivo As String
      motivo = InputBox("Valor abaixo do mínimo!" & vbCrLf & _
                        "Valor Referência: R$ " & Format(ValorRef, "0.00") & vbCrLf & _
                        "Valor Vendido: R$ " & Format(ValorVendido, "0.00") & vbCrLf & vbCrLf & _
                        "Digite o MOTIVO para vender abaixo do mínimo:", _
                        "Autorização Necessária")
      
      If Len(Trim(motivo)) > 0 Then
         ' Gravar log
         vgDb.Execute "INSERT INTO [Vendas Abaixo Minimo] (..." & _
                      "VALUES (" & Now & ", '" & vgPWUsuario & "', ...)"
         PermissaoComLog = True
      Else
         PermissaoComLog = False
      End If
   Else
      PermissaoComLog = False
   End If
End Function
```

---

## Requisitos detalhados para a implementação

- Separação explícita de informações:
  - Valor de Referência (do cadastro) — somente leitura, exibido ao lado.
  - Valor Unitário Orçado (do orçamento) — o valor que realmente entra em total/tributos/validações.
- Regras de preenchimento:
  - Na inclusão de item: preencher o “Valor Orçado” com o “Valor de Referência”.
  - Ao trocar o item: resetar “Valor Orçado” para o “Valor de Referência” do novo item (para não carregar preço do item anterior), com opção de manter se o usuário já editou manualmente — definir critério (ex.: sempre resetar na troca, mais seguro).
- Edição e validação:
  - Usuário pode editar o “Valor Orçado”.
  - Validar “Valor Orçado” >= “Valor de Referência”; se menor, bloquear ou exigir autorização (perfis/fluxo de aprovação).
  - Mensagem clara e uniforme nos três grids.
- UX clara:
  - Mostrar lado a lado: “Ref. Sistema” e “Valor Orçado”.
  - Remover exibição do `Info*` sobre a coluna “Valor Orçado”; usar `ColumnValue` para exibir o que está realmente na linha.
- Impostos e totais:
  - Recalcular impostos e totais sempre que “Valor Orçado” mudar, mantendo compatibilidade com `Entrega_Futura`, NCM, CST/CFOP etc.
- Auditoria:
  - Logar reduções abaixo do mínimo com usuário, motivo e timestamp (quando permitido por permissão).
- Parametrização:
  - Permitir ativar/desativar “respeitar mínimo” por grupo de produtos, cliente ou modalidade.

---

## Proposta técnica (rascunho de abordagem)

1. Colunas: criar/confirmar duas colunas distintas por grid:
   - “Ref. Sistema” (somente leitura, vem de `Info*`)
   - “Valor Unitário” (orçado, lido/escrito via `ColumnValue`)
2. Exibição (CONTEÚDO DA COLUNA):
   - Alterar case da coluna “Valor Unitário” para retornar `Valor_Unitario` (de `ColumnValue`), não mais `Info*`.
   - Adicionar case para “Ref. Sistema” retornando `Info*` (ex.: `InfoProdutos("Valor Unitário")`).
3. Defaults:
   - Em `DEFAULTDASCOLUNAS` ou no momento da seleção do item, definir `ColumnValue(Valor Unitário)` = `Info* ("Valor Unitário")` ao incluir.
   - Ao trocar o item: resetar `ColumnValue(Valor Unitário)` com o novo `Info*` (decidir política se sempre resetar ou só quando vazio/zero).
4. Validações:
   - Manter `ValidaConjx/ValidaPecasx/ValidaProdx` usando o “Valor Orçado”.
   - Exibir mensagem informando o “Ref. Sistema” para clareza, quando inválido.
5. Processamento:
   - Garantir que `ProcessaConjunto/ProcessaPeca/ProcessaProduto` já utilizem o “Valor Orçado” da linha; preservar o fluxo manual quando `Edicao_Manual_Impostos` for verdadeiro.
6. Permissões/Aprovação:
   - Consolidar `Permissao/PermissaoPecas/PermissaoConj` com uma experiência comum (entrada de justificativa, registro em log, etc.).

---

## Riscos e casos de teste

- Casos a cobrir:
  1. Inclusão simples com preço padrão = referência.
  2. Edição para cima (acima do mínimo) — deve aceitar e recalcular impostos.
  3. Edição para baixo (abaixo do mínimo) — deve bloquear ou pedir permissão e registrar.
  4. Troca de item — deve resetar o “Valor Orçado” para o novo “Ref. Sistema”.
  5. Entrega futura (CFOP/CST e impostos específicos) — conferir que impostos recalculam corretamente após alteração de preço.
  6. NCM inativo — continuar bloqueando conforme hoje em `ProcessaProduto`.
  7. `Edicao_Manual_Impostos = True` — respeitar fluxo manual em Produtos.
  8. Desconto/Frete alterados — recalcular PIS/COFINS e ICMS com base no novo total.
  9. Produtos grupo 20 — garantir que o “mínimo do sistema” exibido e o validado sejam a mesma regra (custo*3,5 ou outro critério aprovado), e que a validação acompanhe a exibição.

Notas adicionais do fluxo:
- Após inclusão/edição/exclusão nos grids, o sistema chama `TotalizaOrcamento` para atualizar totais do orçamento.
- Defaults fiscais: em `DEFAULTDASCOLUNAS`, CST/CFOP são preenchidos com base em `Fatura_Proforma` (ex.: CST 41, CFOP 7101) — comportamento presente nos três grids.
- Permissões de edição: mesmo com restrições (pedido/cancelado/venda fechada), `Edicao_Manual_Impostos` permite edição em grids; em alguns casos há exceções por usuário (YGOR, JUCELI, MAYSA).

---

## Próximos passos (decisões e implementação)

1. **Decisão crítica — Produtos grupo 20**: 
   - Definir qual é o "Valor de Referência" oficial: `ProdutoAux![Valor Total]` (usado na validação hoje) ou `Valor de Custo * 3.5` (usado na exibição quando MP 43602)?
   - Proposta: **unificar usando a mesma regra** (provavelmente custo*3.5 quando aplicável) tanto em `InfoProdutos` quanto em `ValidaProdx`, para que exibição e validação sejam coerentes.

2. **Implementação da abordagem de duas colunas**:
   - Adicionar coluna "Ref. Sistema" (somente leitura) em cada grid antes da coluna "Valor Unitário".
   - Alterar `CONTEUDODACOLUNA` da coluna "Valor Unitário" para retornar `Valor_Unitario` (de `ColumnValue`), não mais `Info*`.
   - A nova coluna "Ref. Sistema" mostra `Info*("Valor Unitário")` para referência visual.

3. **Defaults e reset ao trocar item**:
   - Em `DEFAULTDASCOLUNAS`, preencher "Valor Unitário" com `Info*("Valor Unitário")` na inclusão.
   - Detectar troca de item em `IniApDaCol` (comparar sequência anterior com atual): se trocou, resetar "Valor Unitário" = `Info*` do novo item.
   - Alternativa: sempre usar o valor do `ColumnValue` e deixar o usuário ajustar manualmente após selecionar o item (mais simples, mas pode gerar esquecimentos).

4. **Validações consistentes**:
   - Manter `ValidaConjx/ValidaPecasx/ValidaProdx` comparando com o valor de referência correto.
   - Melhorar mensagem de erro para incluir o valor mínimo: "Valor Unitário não pode ser menor que R$ [valor_ref] (Valor de Referência do Sistema)".

5. **Implementação em fases**:
   - **Fase 1**: Produtos (Grid 3) — mais complexo, tem exceção grupo 20, modo manual e mais impostos. Resolver aqui primeiro garante que as outras duas fases serão mais simples.
   - **Fase 2**: Peças (Grid 1) — similar a Produtos mas sem exceções de grupo.
   - **Fase 3**: Conjuntos (Grid 0) — mais direto, sem exceções especiais.

6. **Testes após cada fase**:
   - Inclusão com preço padrão.
   - Edição para valor acima do mínimo.
   - Tentativa de edição abaixo do mínimo (com e sem permissão).
   - Troca de item (confirmar reset do preço).
   - Entrega futura e impostos.
   - NCM inativo (só Produtos).
   - Modo manual (`Edicao_Manual_Impostos = True`).

7. **Auditoria (opcional, mas recomendado)**:
   - Criar tabela de log para registrar quando usuário usa permissão para vender abaixo do mínimo.
   - Campos: data/hora, usuário, orçamento, item, valor mínimo, valor vendido, motivo (texto livre).

---

## Decisão recomendada para começar

**Começar por Produtos (Grid 3), resolvendo primeiro a inconsistência grupo 20:**

1. Alinhar `InfoProdutos("Valor Unitário")` e `ValidaProdx` para usar a **mesma regra** (custo*3.5 quando aplicável).
2. Adicionar coluna "Ref. Sistema" antes da coluna 10 (Valor Unitário).
3. Mudar CONTEUDODACOLUNA col 10 para retornar `Valor_Unitario` de `ColumnValue`.
4. Em DEFAULTDASCOLUNAS col 10, preencher com `InfoProdutos("Valor Unitário")` na inclusão.
5. Testar todos os cenários listados acima.
6. Após validado em Produtos, replicar para Peças e Conjuntos (que são mais simples).

---

## Apêndice — Pontos de código relevantes

- `InfoProdutos/InfoPecas/InfoConjuntos`: retornam NCM, Sigla, Valor Unitário (referência), Estoque, Peso etc. — hoje usados também para renderizar a coluna de Valor Unitário.
- `ExecutaGrid0/1/3`:
  - CONTEÚDO DA COLUNA das respectivas colunas de Valor Unitário chamadas:
    - Grid(0) col 7: `InfoConjuntos("Valor Unitário")`
    - Grid(1) col 9: `InfoPecas("Valor Unitário")`
    - Grid(3) col 10: `InfoProdutos("Valor Unitário")`
  - Totais usam `Quantidade * Valor_Unitario` (variável vinda de `ColumnValue`).
- Validações:
  - `ValidaConjx/ValidaPecasx/ValidaProdx` comparam o Valor Orçado (da linha) contra o Valor de Referência (do cadastro) e permitem exceções por permissão.
- Processamento:
  - `ProcessaConjunto/ProcessaPeca/ProcessaProduto` gravam valores e recalculam tributos. Em Produtos, há caminho manual quando `Edicao_Manual_Impostos = True`.

---

## Checklist final — O que validar antes de dar como pronto

### Funcionalidade básica (todos os grids)
- [ ] Coluna "Ref. Sistema" exibe valor do cadastro corretamente
- [ ] Coluna "Valor Unitário" exibe valor do orçamento (ColumnValue)
- [ ] Total = Quantidade × Valor Unitário (do orçamento, não do cadastro)
- [ ] Na inclusão de novo item, "Valor Unitário" = "Ref. Sistema"
- [ ] Ao trocar item, "Valor Unitário" reseta para "Ref. Sistema" do novo item
- [ ] Edição para valor acima do mínimo funciona normalmente
- [ ] Edição para valor abaixo do mínimo bloqueia ou pede permissão
- [ ] Mensagem de erro informa claramente o valor mínimo
- [ ] Impostos recalculam após alteração de "Valor Unitário"

### Casos específicos
- [ ] Produtos grupo 20: exibição e validação usam mesma regra (custo*3.5 ou Valor Total)
- [ ] Produtos sem matéria-prima: comportamento documentado e correto
- [ ] Entrega futura: CFOP/CST/impostos específicos funcionam
- [ ] NCM inativo: continua bloqueando em ProcessaProduto
- [ ] Modo manual (Edicao_Manual_Impostos): funciona conforme esperado
- [ ] Desconto/Frete: recalculam PIS/COFINS/ICMS corretamente
- [ ] Usuários com permissão: lista funciona ou foi migrada para config
- [ ] Log/auditoria: vendas abaixo do mínimo são registradas (se implementado)

### Integração e regressão
- [ ] TotalizaOrcamento atualiza totais do orçamento corretamente
- [ ] Defaults CST/CFOP (Fatura_Proforma) continuam funcionando
- [ ] Permissões de edição (pedido/cancelado/venda fechada) respeitadas
- [ ] Navegação entre registros não causa comportamentos inesperados
- [ ] Salvar/cancelar operações no grid funcionam normalmente
- [ ] Relatórios e consultas que usam "Valor Unitário" continuam corretos

---

## Resumo — O que mudou e por quê

**ANTES (estado problemático)**:
- Coluna "Valor Unitário" mostrava valor do cadastro (`Info*`).
- Cálculos usavam valor do orçamento (`ColumnValue`).
- Usuário via um número, sistema trabalhava com outro.
- Troca de item mantinha preço do item anterior.
- Produtos grupo 20: exibição e validação divergentes.

**DEPOIS (estado corrigido)**:
- Coluna "Ref. Sistema" (nova) mostra valor do cadastro — somente leitura.
- Coluna "Valor Unitário" mostra valor do orçamento — editável.
- Usuário vê exatamente o que o sistema usa.
- Troca de item reseta automaticamente o preço.
- Produtos grupo 20: exibição e validação alinhadas.
- Mensagens claras informando valor mínimo.
- Auditoria de exceções (opcional).

**BENEFÍCIOS**:
- ✅ Transparência total para o usuário
- ✅ Elimina bugs de preço "preso" de item anterior
- ✅ Facilita manutenção do código
- ✅ Permite rastreabilidade de vendas especiais
- ✅ Base sólida para futuras melhorias (ex.: tabelas de preço, descontos progressivos, etc.)

---

## Próxima ação recomendada

**VOCÊ PRECISA DECIDIR AGORA** (antes de começar a implementar):

1. **Produtos grupo 20 — Qual regra de "Valor de Referência" usar?**
   - [ ] Opção A: Manter `Custo * 3.5` e alinhar validação com exibição
   - [ ] Opção B: Remover regra especial e sempre usar `Valor Total`
   - [ ] Opção C: Outra (especificar): _________________________

2. **Reset ao trocar item — Como funciona?**
   - [ ] Opção A: Sempre resetar "Valor Unitário" = "Ref. Sistema"
   - [ ] Opção B: Só resetar se "Valor Unitário" estiver vazio/zero
   - [ ] Opção C: Perguntar ao usuário (popup)

3. **Adicionar coluna no grid — Formato?**
   - [ ] Opção A: Adicionar coluna física "Ref. Sistema"
   - [ ] Opção B: Mostrar "Ref. Sistema" em tooltip/hint

4. **Auditoria — Quando implementar?**
   - [ ] Opção A: Junto com a correção (fase 4)
   - [ ] Opção B: Em versão futura (não agora)

**Após decidir, começamos pela Fase 1 (Grid 3 - Produtos) e validamos cada passo antes de replicar para Peças e Conjuntos.**
