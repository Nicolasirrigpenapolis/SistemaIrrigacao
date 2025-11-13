# Implementação Opção 1 - Fix: Troca de item mantém preço errado

**Data**: 12/11/2025  
**Solução**: Opção 1 - Detecção manual de troca de item

---

## Por que Opção 2 não funcionou

`DEFAULTDASCOLUNAS` não é chamado automaticamente quando o usuário troca o item na coluna 1 de uma linha já existente. Ele só é chamado em:
- Nova linha (ACAO_INCLUINDO)
- Alguns eventos específicos do grid

Por isso, precisamos da **Opção 1**: detecção manual na função `IniApDaCol`.

---

## O que foi implementado

### 1. Variáveis de controle (Linha ~10105)

Adicionadas 3 variáveis no nível do Form para rastrear a sequência anterior:

```vb
'Variáveis para detectar troca de item (Fix: preço errado ao trocar item)
Dim Seq_Conjunto_Anterior As Long
Dim Seq_Produto_Pecas_Anterior As Long
Dim Seq_Produto_Anterior As Long
```

### 2. Grid(0) - Conjuntos (Linha ~20795)

Modificado `IniApDaCol` para detectar troca e resetar:

```vb
IniApDaCol:
   On Error Resume Next
   Dim Sequencia_Nova As Long
   Sequencia_Nova = val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   
   ' Detecta troca de item e reseta valor unitário
   If Grid(0).Status = ACAO_EDITANDO And Sequencia_Nova <> Seq_Conjunto_Anterior And Seq_Conjunto_Anterior <> 0 And Sequencia_Nova > 0 Then
      ' RESET: busca valor do cadastro do novo item
      Set ConjuntoAux = Conjuntos.Seek("=", Sequencia_Nova)
      If Not ConjuntoAux.EOF Then
         ColumnValue(7) = ConjuntoAux![Valor Total]
      End If
   End If
   
   ' Atualiza tracker de sequência
   Seq_Conjunto_Anterior = Sequencia_Nova
   
   Sequencia_do_Conjunto = Sequencia_Nova
   ...resto do código...
```

### 3. Grid(1) - Peças (Linha ~21126)

Modificado `IniApDaCol`:

```vb
IniApDaCol:
   On Error Resume Next
   Dim Sequencia_Nova As Long
   Sequencia_Nova = val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   
   ' Detecta troca de item e reseta valor unitário
   If Grid(1).Status = ACAO_EDITANDO And Sequencia_Nova <> Seq_Produto_Pecas_Anterior And Seq_Produto_Pecas_Anterior <> 0 And Sequencia_Nova > 0 Then
      ' RESET: busca valor do cadastro do novo item
      Set ProdutoAux = Produtos.Seek("=", Sequencia_Nova)
      If Not ProdutoAux.EOF Then
         ColumnValue(9) = ProdutoAux![Valor Total]
      End If
   End If
   
   ' Atualiza tracker de sequência
   Seq_Produto_Pecas_Anterior = Sequencia_Nova
   
   Sequencia_do_Produto = Sequencia_Nova
   ...resto do código...
```

### 4. Grid(3) - Produtos (Linha ~21605)

Modificado `IniApDaCol` com regra especial para grupo 20:

```vb
IniApDaCol:
   On Error Resume Next
   Dim Sequencia_Nova As Long
   Sequencia_Nova = val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   
   ' Detecta troca de item e reseta valor unitário
   If Grid(3).Status = ACAO_EDITANDO And Sequencia_Nova <> Seq_Produto_Anterior And Seq_Produto_Anterior <> 0 And Sequencia_Nova > 0 Then
      ' RESET: busca valor do cadastro do novo item
      Set ProdutoAux = Produtos.Seek("=", Sequencia_Nova)
      If Not ProdutoAux.EOF Then
         ' Aplica regra especial para grupo 20
         If ProdutoAux![Seqüência do Grupo de Produto] = 20 Then
            ' Verifica se tem MP 43602 no orçamento
            Dim rs_MP As GRecordSet
            Set rs_MP = Produtos_do_Orcamento.Filter("[Seqüência do Produto] = 43602")
            If rs_MP.RecordCount > 0 Then
               ColumnValue(10) = ProdutoAux![Valor de Custo] * 3.5
            Else
               ColumnValue(10) = ProdutoAux![Valor Total]
            End If
         Else
            ColumnValue(10) = ProdutoAux![Valor Total]
         End If
      End If
   End If
   
   ' Atualiza tracker de sequência
   Seq_Produto_Anterior = Sequencia_Nova
   
   Sequencia_do_Produto = Sequencia_Nova
   ...resto do código...
```

---

## Como funciona

### Detecção de troca

A cada chamada de `IniApDaCol`:
1. Lê a `Sequencia_Nova` de `ColumnValue(1)`
2. Compara com `Seq_*_Anterior`
3. Se **diferente** E `Status = ACAO_EDITANDO` E `Anterior <> 0` E `Nova > 0` → **detectou troca**

### Condições para reset

```vb
Grid(X).Status = ACAO_EDITANDO           ' Não é nova linha
Sequencia_Nova <> Seq_*_Anterior         ' Sequência mudou
Seq_*_Anterior <> 0                      ' Não é primeira vez
Sequencia_Nova > 0                       ' Nova sequência válida
```

### Reset do valor

1. Busca o registro no cadastro: `Produtos.Seek("=", Sequencia_Nova)`
2. Se encontrou:
   - **Conjuntos**: `ColumnValue(7) = ConjuntoAux![Valor Total]`
   - **Peças**: `ColumnValue(9) = ProdutoAux![Valor Total]`
   - **Produtos**: 
     - Se grupo 20 + MP 43602: `ColumnValue(10) = Custo * 3.5`
     - Senão: `ColumnValue(10) = Valor Total`

### Atualização do tracker

Sempre atualiza a variável de controle:
```vb
Seq_*_Anterior = Sequencia_Nova
```

---

## O que isso resolve

✅ **Problema #2**: Troca de item mantém preço errado  
- Antes: Ao trocar Produto A → B, `ColumnValue` mantinha preço do A
- Agora: Ao trocar, `ColumnValue(10)` é forçado para o preço do B

✅ **Mantém comportamento correto**:
- Edição manual continua funcionando (só reseta na TROCA, não na edição)
- Validação de mínimo continua bloqueando valores abaixo do permitido
- Produtos grupo 20 aplicam regra `custo*3.5` quando tem MP 43602
- Nova linha inicia com `Seq_*_Anterior = 0`, então não reseta na primeira vez

---

## Testes obrigatórios

### ✅ Teste 1: Troca simples de item
**Passos**:
1. Grid Produtos: adicionar Produto A (R$ 100)
2. Na mesma linha, trocar coluna 1 para Produto B (R$ 200)
3. Verificar Valor Unitário e total

**Resultado esperado**:
- Valor Unitário mostra R$ 200
- Total calcula com R$ 200

---

### ✅ Teste 2: Troca após edição manual
**Passos**:
1. Grid Produtos: adicionar Produto A (R$ 100)
2. Editar Valor Unitário para R$ 95
3. Trocar coluna 1 para Produto B (R$ 200)

**Resultado esperado**:
- Valor Unitário reseta para R$ 200 (não mantém R$ 95)

---

### ✅ Teste 3: Edição sem troca
**Passos**:
1. Grid Produtos: adicionar Produto A (R$ 100)
2. Editar Valor Unitário para R$ 95
3. Sair da célula e voltar

**Resultado esperado**:
- Valor mantém R$ 95 (edição preservada)
- ⚠️ **IMPORTANTE**: Não deve resetar!

---

### ✅ Teste 4: Nova linha
**Passos**:
1. Adicionar nova linha
2. Selecionar Produto C (R$ 150)

**Resultado esperado**:
- Valor exibe R$ 150 (via Info*)
- ColumnValue pode estar vazio ou com valor anterior
- Na troca para outro produto, reseta corretamente

---

### ✅ Teste 5: Validação mínimo
**Passos**:
1. Adicionar Produto A (R$ 100)
2. Editar para R$ 80
3. Tentar gravar

**Resultado esperado**:
- Sistema bloqueia ou pede permissão

---

### ✅ Teste 6: Produtos grupo 20
**Passos**:
1. Adicionar produto grupo 20
2. Verificar se tem MP 43602
3. Trocar para outro produto grupo 20

**Resultado esperado**:
- Com MP 43602: `Valor = Custo * 3.5`
- Sem MP 43602: `Valor = Valor Total`

---

### ✅ Teste 7: Múltiplas trocas
**Passos**:
1. Adicionar Produto A (R$ 100)
2. Trocar para Produto B (R$ 200)
3. Trocar para Produto C (R$ 150)
4. Verificar valor a cada troca

**Resultado esperado**:
- Cada troca reseta para o preço correto do novo produto

---

## Arquivos modificados

- `IRRIG/ORCAMENT.FRM`
  - Declarações (linha ~10105): 3 variáveis adicionadas
  - Grid(0) IniApDaCol (linha ~20795): lógica de reset
  - Grid(1) IniApDaCol (linha ~21126): lógica de reset
  - Grid(3) IniApDaCol (linha ~21605): lógica de reset + regra grupo 20

---

## Próximos passos

1. ✅ Compilar o projeto VB6
2. ⏸️ Executar Testes 1-7 em CADA grid
3. ⏸️ Se tudo OK → Commit
4. ⏸️ Monitorar em produção

---

## Notas técnicas

### Por que `Seq_*_Anterior <> 0`?

Quando a linha é NOVA (primeira vez em `IniApDaCol`), a variável está zerada. Isso evita que tente resetar na primeira leitura.

### Por que `Grid(X).Status = ACAO_EDITANDO`?

Para garantir que é uma linha JÁ EXISTENTE sendo editada, não uma nova linha sendo incluída.

### Por que `Sequencia_Nova > 0`?

Para garantir que o novo item é válido antes de buscar no cadastro.

### E se o Seek falhar?

O `If Not ProdutoAux.EOF` protege contra itens não encontrados. Se o seek falhar, simplesmente não reseta o valor.
