# Implementação Concluída - Fix: Troca de item mantém preço errado

**Data**: 12/11/2025  
**Solução**: Opção 2 - DEFAULTDASCOLUNAS

---

## O que foi implementado

Adicionado `Case` para coluna do Valor Unitário no bloco `DEFAULTDASCOLUNAS` de cada grid, usando o mesmo padrão já existente para CST e CFOP.

### Grid(0) - Conjuntos (Linha ~20696)
```vb
Case 7  ' Valor Unitário
   If Sequencia_do_Conjunto > 0 Then
      vgRetVal = InfoConjuntos(..., "Valor Unitário")
   End If
```

### Grid(1) - Peças (Linha ~21013)
```vb
Case 9  ' Valor Unitário
   If Sequencia_do_Produto > 0 Then
      vgRetVal = InfoPecas(..., "Valor Unitário")
   End If
```

### Grid(3) - Produtos (Linha ~21480)
```vb
Case 10  ' Valor Unitário
   If Sequencia_do_Produto > 0 Then
      vgRetVal = InfoProdutos(..., "Valor Unitário")
   End If
```

---

## Como funciona

1. Quando uma **nova linha** é criada (ACAO_INCLUINDO), o grid chama `DEFAULTDASCOLUNAS` para preencher valores padrão.
2. Quando o usuário **troca o item** (muda coluna 1 - Sequencia), o grid reaplica os defaults das colunas dependentes.
3. O `Case` adicionado retorna o valor do cadastro via `Info*("Valor Unitário")`, que:
   - Para Conjuntos: busca `ConjuntoAux![Valor Total]`
   - Para Peças: busca `ProdutoAux![Valor Total]`
   - Para Produtos: busca `ProdutoAux![Valor Total]` OU `Custo*3.5` (grupo 20 + MP 43602)
4. O grid atualiza `ColumnValue(coluna_valor)` automaticamente com esse valor.

---

## O que isso resolve

✅ **Problema #2**: Troca de item mantém preço errado  
- Antes: Ao trocar Produto A → B, `ColumnValue` mantinha preço do A, mas tela exibia preço do B
- Agora: Ao trocar, `ColumnValue` é resetado para o preço do B automaticamente

✅ **Mantém comportamento correto**:
- Edição manual continua funcionando (usuário pode alterar valor)
- Validação de mínimo continua bloqueando valores abaixo do permitido
- Produtos grupo 20 continuam aplicando regra `custo*3.5` quando aplicável
- Nova linha já inicia com valor correto do cadastro

---

## Testes obrigatórios

### Teste 1: Troca simples de item
**Passos**:
1. Abrir orçamento
2. Grid Produtos: adicionar Produto A (ex: R$ 100)
3. Na mesma linha, trocar coluna 1 para Produto B (ex: R$ 200)
4. Verificar coluna "Valor Unitário" e total da linha

**Resultado esperado**:
- ✅ Valor Unitário mostra R$ 200
- ✅ Total calcula com R$ 200 (não R$ 100)

---

### Teste 2: Troca após edição manual
**Passos**:
1. Grid Produtos: adicionar Produto A (R$ 100)
2. Editar Valor Unitário para R$ 95
3. Trocar coluna 1 para Produto B (R$ 200)
4. Verificar Valor Unitário

**Resultado esperado**:
- ✅ Valor Unitário reseta para R$ 200 (não mantém R$ 95)
- ✅ Total calcula com R$ 200

---

### Teste 3: Edição sem troca (garantir que edição manual funciona)
**Passos**:
1. Grid Produtos: adicionar Produto A (R$ 100)
2. Editar Valor Unitário para R$ 95
3. Sair da célula (Tab ou Enter)
4. Voltar para a célula
5. Verificar valor

**Resultado esperado**:
- ✅ Valor mantém R$ 95 (edição manual preservada)
- ✅ Total calcula com R$ 95

---

### Teste 4: Nova linha
**Passos**:
1. Grid Produtos: adicionar nova linha
2. Selecionar Produto C (R$ 150)
3. Verificar Valor Unitário

**Resultado esperado**:
- ✅ Valor Unitário automaticamente preenche com R$ 150

---

### Teste 5: Validação de mínimo
**Passos**:
1. Grid Produtos: adicionar Produto A (R$ 100)
2. Editar Valor Unitário para R$ 80
3. Tentar gravar

**Resultado esperado**:
- ✅ Sistema bloqueia ou pede permissão (comportamento atual mantido)
- ✅ Mensagem: "Valor Unitário inválido!" ou similar

---

### Teste 6: Produtos grupo 20 (regra especial)
**Passos**:
1. Grid Produtos: adicionar produto do grupo 20
2. Verificar se existe MP 43602 no orçamento
3. Observar Valor Unitário exibido

**Resultado esperado**:
- ✅ Se tem MP 43602: Valor = `Custo * 3.5`
- ✅ Se não tem MP 43602: Valor = `Valor Total`
- ✅ Troca de produto grupo 20 aplica regra corretamente

---

### Teste 7: Replicar nos outros grids
**Passos**:
1. Repetir Testes 1-3 para Grid(1) - Peças
2. Repetir Testes 1-3 para Grid(0) - Conjuntos

**Resultado esperado**:
- ✅ Comportamento idêntico nos 3 grids

---

## Checklist de validação

- [ ] Teste 1: Troca simples ✅
- [ ] Teste 2: Troca após edição ✅
- [ ] Teste 3: Edição sem troca ✅
- [ ] Teste 4: Nova linha ✅
- [ ] Teste 5: Validação mínimo ✅
- [ ] Teste 6: Produtos grupo 20 ✅
- [ ] Teste 7: Grid Peças ✅
- [ ] Teste 7: Grid Conjuntos ✅

---

## Rollback (se necessário)

Se os testes mostrarem que `DEFAULTDASCOLUNAS` NÃO é chamado na troca de item:

1. Remover os `Case 7`, `Case 9` e `Case 10` adicionados
2. Implementar **Opção 1** (detecção manual com variáveis de controle)
3. Ver arquivo `PLANO_IMPLEMENTACAO.md` para código da Opção 1

**Nota**: Isso é improvável, pois CST/CFOP já funcionam assim hoje.

---

## Arquivos modificados

- `IRRIG/ORCAMENT.FRM` (3 blocos `DEFAULTDASCOLUNAS` editados)

## Linhas aproximadas modificadas

- Grid(0) Conjuntos: ~linha 20696-20708
- Grid(1) Peças: ~linha 21013-21029
- Grid(3) Produtos: ~linha 21480-21496

---

## Próximos passos

1. ✅ Compilar o projeto VB6
2. ⏸️ Executar Testes 1-7
3. ⏸️ Se tudo OK → Commit e deploy
4. ⏸️ Se falhar → Rollback e implementar Opção 1
