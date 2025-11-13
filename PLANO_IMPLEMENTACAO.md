# Plano de Implementação - Fix: Troca de item mantém preço errado

## Problema identificado

Quando o usuário:
1. Seleciona um item (Produto A com preço R$ 100)
2. Edita o valor unitário para R$ 95
3. Troca para outro item (Produto B com preço R$ 200)

**O que acontece**:
- `ColumnValue(coluna_do_valor)` mantém R$ 95 (valor editado do item anterior)
- Tela exibe R$ 200 via `Info*` (valor do cadastro do novo item)
- Sistema calcula com R$ 95 (ERRADO!)

## Causa raiz

Na função `IniApDaCol` de cada grid, quando troca de item:
- A coluna 1 (Sequencia_do_*) é atualizada com o novo item
- **MAS** `ColumnValue(coluna_valor_unitario)` NÃO é resetado
- Ele mantém o valor editado anteriormente

**Locais identificados**:

### Grid(0) - Conjuntos (linha ~20788)
```
Sequencia_do_Conjunto = val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
...
Valor_Unitario = ColumnValue(7)  ← mantém valor antigo!
```

### Grid(1) - Peças (linha ~21119)
```
Sequencia_do_Produto = val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
...
Valor_Unitario = ColumnValue(9)  ← mantém valor antigo!
```

### Grid(3) - Produtos (linha ~21598)
```
Sequencia_do_Produto = val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
...
Valor_Unitario = ColumnValue(10)  ← mantém valor antigo!
```

## Solução proposta

### Detectar quando houve troca de item

Precisamos saber se o `Sequencia_do_*` mudou em relação ao registro anterior.

**Estratégia**:
1. Guardar a `Sequencia` anterior em variável de módulo/form
2. Na `IniApDaCol`, comparar `Sequencia` atual com a anterior
3. Se DIFERENTE → houve troca de item → RESETAR `ColumnValue` do valor unitário

### Código a adicionar

**No nível do Form** (declarações no topo):
```vb
' Variáveis para detectar troca de item
Dim Seq_Conjunto_Anterior As Long
Dim Seq_Produto_Pecas_Anterior As Long
Dim Seq_Produto_Anterior As Long
```

**Grid(0) - Conjuntos** (na `IniApDaCol`, após linha ~20790):
```vb
IniApDaCol:
   On Error Resume Next
   Dim Sequencia_Nova As Long
   Sequencia_Nova = val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   
   ' Detecta troca de item
   If Grid(0).Status = ACAO_EDITANDO And Sequencia_Nova <> Seq_Conjunto_Anterior And Seq_Conjunto_Anterior <> 0 Then
      ' RESET: busca valor do cadastro do novo item
      If Sequencia_Nova > 0 Then
         Set ConjuntoAux = Conjuntos.Seek("=", Sequencia_Nova)
         If Not ConjuntoAux.EOF Then
            ColumnValue(7) = ConjuntoAux![Valor Total]  ' Reset valor unitário
         End If
      End If
   End If
   
   Seq_Conjunto_Anterior = Sequencia_Nova  ' Atualiza tracker
   
   Sequencia_do_Conjunto = Sequencia_Nova
   CST = ColumnValue(2)
   ...continua normal...
```

**Grid(1) - Peças** (na `IniApDaCol`, após linha ~21121):
```vb
IniApDaCol:
   On Error Resume Next
   Dim Sequencia_Nova As Long
   Sequencia_Nova = val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   
   ' Detecta troca de item
   If Grid(1).Status = ACAO_EDITANDO And Sequencia_Nova <> Seq_Produto_Pecas_Anterior And Seq_Produto_Pecas_Anterior <> 0 Then
      ' RESET: busca valor do cadastro do novo item
      If Sequencia_Nova > 0 Then
         Set ProdutoAux = Produtos.Seek("=", Sequencia_Nova)
         If Not ProdutoAux.EOF Then
            ColumnValue(9) = ProdutoAux![Valor Total]  ' Reset valor unitário
         End If
      End If
   End If
   
   Seq_Produto_Pecas_Anterior = Sequencia_Nova  ' Atualiza tracker
   
   Sequencia_do_Produto = Sequencia_Nova
   CST = ColumnValue(2)
   ...continua normal...
```

**Grid(3) - Produtos** (na `IniApDaCol`, após linha ~21600):
```vb
IniApDaCol:
   On Error Resume Next
   Dim Sequencia_Nova As Long
   Sequencia_Nova = val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   
   ' Detecta troca de item
   If Grid(3).Status = ACAO_EDITANDO And Sequencia_Nova <> Seq_Produto_Anterior And Seq_Produto_Anterior <> 0 Then
      ' RESET: busca valor do cadastro do novo item
      If Sequencia_Nova > 0 Then
         Set ProdutoAux = Produtos.Seek("=", Sequencia_Nova)
         If Not ProdutoAux.EOF Then
            ' Regra especial: grupo 20 + MP 43602
            If Not ProdutoAux.EOF Then
               If ProdutoAux![Seqüência do Grupo de Produto] = 20 Then
                  ' Verifica se tem MP 43602
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
      End If
   End If
   
   Seq_Produto_Anterior = Sequencia_Nova  ' Atualiza tracker
   
   Sequencia_do_Produto = Sequencia_Nova
   CST = ColumnValue(3)
   ...continua normal...
```

## Cuidados importantes

### ✅ O que a solução FAZ

1. **Detecta troca de item**: Compara sequência atual com anterior
2. **Reset automático**: Quando detecta troca, busca valor do cadastro do NOVO item
3. **Permite edição manual**: Usuário pode editar normalmente, não afeta
4. **Mantém validação existente**: `Valida*x` continua bloqueando valores abaixo do mínimo

### ✅ O que a solução NÃO AFETA

1. **Edição normal do valor**: Se usuário editar valor SEM trocar item, funciona normal
2. **Validação de mínimo**: `Valida*x` continua funcionando como sempre
3. **Permissões**: Usuários autorizados ainda podem vender abaixo do mínimo
4. **Inclusão de novo item**: Não afeta ACAO_INCLUINDO (nova linha)

## Testes necessários

1. **Teste 1 - Troca simples**:
   - Adicionar Produto A, não editar valor
   - Trocar para Produto B
   - ✅ Verificar: valor reseta para preço do Produto B

2. **Teste 2 - Troca com edição**:
   - Adicionar Produto A (R$ 100), editar para R$ 95
   - Trocar para Produto B (R$ 200)
   - ✅ Verificar: valor reseta para R$ 200, cálculo usa R$ 200

3. **Teste 3 - Edição sem troca**:
   - Adicionar Produto A (R$ 100)
   - Editar valor para R$ 95
   - Sair da linha e voltar
   - ✅ Verificar: valor mantém R$ 95 (não reseta!)

4. **Teste 4 - Produtos grupo 20**:
   - Adicionar produto grupo 20 com MP 43602 presente
   - Trocar para outro produto grupo 20
   - ✅ Verificar: aplica regra custo*3.5 se aplicável

5. **Teste 5 - Validação de mínimo**:
   - Adicionar Produto A (R$ 100)
   - Editar para R$ 80
   - ✅ Verificar: validação bloqueia ou pede permissão (comportamento atual mantido)

## Ordem de implementação

1. **Fase 1**: Adicionar declarações das variáveis de controle no Form
2. **Fase 2**: Implementar fix no Grid(3) - Produtos (mais complexo, tem regra grupo 20)
3. **Fase 3**: Testar completamente Grid(3)
4. **Fase 4**: Implementar fix no Grid(1) - Peças
5. **Fase 5**: Testar completamente Grid(1)
6. **Fase 6**: Implementar fix no Grid(0) - Conjuntos
7. **Fase 7**: Teste final integrado

**Importante**: NÃO prosseguir para próxima fase sem testar completamente a anterior!
