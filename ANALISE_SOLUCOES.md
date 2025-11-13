# Análise: Melhor forma de resolver "Troca de item mantém preço errado"

## Problema

Quando usuário troca de item (muda coluna 1 - Sequencia), o `ColumnValue` do Valor Unitário mantém o valor antigo.

---

## OPÇÃO 1: Detectar troca em IniApDaCol (Solução proposta inicial)

### Como funciona
- Adicionar variáveis de controle no Form (`Seq_*_Anterior`)
- Na função `IniApDaCol`, comparar sequência atual com anterior
- Se diferente → buscar valor do cadastro e forçar `ColumnValue = valor_cadastro`

### ✅ Vantagens
- Lógica centralizada em um lugar só
- Funciona em qualquer situação que chame `IniApDaCol`
- Controle total sobre quando resetar

### ❌ Desvantagens
- Precisa criar 3 variáveis de controle no Form
- Lógica "manual" de detecção de mudança
- Precisa manter estado (variável anterior)
- Se esquecer de atualizar a variável, quebra
- **Mais código** e mais complexo

### Código necessário
```vb
' No Form (declarações)
Dim Seq_Conjunto_Anterior As Long
Dim Seq_Produto_Pecas_Anterior As Long
Dim Seq_Produto_Anterior As Long

' Em cada IniApDaCol (3 lugares)
IniApDaCol:
   On Error Resume Next
   Dim Sequencia_Nova As Long
   Sequencia_Nova = val(Parse$(CStr(ColumnValue(1)), Chr$(1), 1))
   
   If Grid(X).Status = ACAO_EDITANDO And Sequencia_Nova <> Seq_*_Anterior And Seq_*_Anterior <> 0 Then
      If Sequencia_Nova > 0 Then
         Set *Aux = *.Seek("=", Sequencia_Nova)
         If Not *Aux.EOF Then
            ColumnValue(N) = *Aux![Valor Total]
         End If
      End If
   End If
   
   Seq_*_Anterior = Sequencia_Nova
   ...resto do código...
```

---

## OPÇÃO 2: Usar DEFAULTDASCOLUNAS para Valor Unitário ⭐ MELHOR

### Como funciona
- O sistema JÁ tem um mecanismo: `DEFAULTDASCOLUNAS`
- É chamado automaticamente quando:
  - Nova linha é criada (ACAO_INCLUINDO)
  - **Coluna muda e precisa de valor default**
- Adicionar `Case` para coluna do Valor Unitário em `DEFAULTDASCOLUNAS`
- Retornar valor do cadastro (via `Info*`)

### ✅ Vantagens
- **Usa mecanismo nativo do sistema** (já existe!)
- **Muito menos código** (apenas adicionar Case)
- **Não precisa variáveis de controle**
- **Não precisa rastrear estado**
- **Mais limpo e manutenível**
- Se funciona para CST e CFOP, vai funcionar para Valor Unitário

### ❌ Desvantagens
- Precisa entender QUANDO o sistema chama `DEFAULTDASCOLUNAS`
- Se o sistema não chamar na troca de item, não funciona
- Menos controle explícito

### Código necessário
```vb
' Grid(0) - Conjuntos (linha ~20696)
ElseIf vgOq = DEFAULTDASCOLUNAS Then
   GoSub IniApDaCol
   vgRetVal = Null
   Select Case vgCol
      Case 2
         CST = IIf(Fatura_Proforma, 41, 0)
         vgRetVal = CST
      Case 3
         CFOP = IIf(Fatura_Proforma, 7101, 0)
         vgRetVal = CFOP
      Case 7  ' ← NOVO: Valor Unitário
         If Sequencia_do_Conjunto > 0 Then
            vgRetVal = InfoConjuntos(..., "Valor Unitário")
         End If
   End Select

' Grid(1) - Peças (linha ~21013)
ElseIf vgOq = DEFAULTDASCOLUNAS Then
   GoSub IniApDaCol
   vgRetVal = Null
   Select Case vgCol
      Case 3
         CST = IIf(Fatura_Proforma, 41, 0)
         vgRetVal = CST
      Case 4
         CFOP = IIf(Fatura_Proforma, 7101, 0)
         vgRetVal = CFOP
      Case 9  ' ← NOVO: Valor Unitário
         If Sequencia_do_Produto > 0 Then
            vgRetVal = InfoPecas(..., "Valor Unitário")
         End If
   End Select

' Grid(3) - Produtos (linha ~21480)
ElseIf vgOq = DEFAULTDASCOLUNAS Then
   GoSub IniApDaCol
   vgRetVal = Null
   Select Case vgCol
      Case 3
         CST = IIf(Fatura_Proforma, 41, 0)
         vgRetVal = CST
      Case 4
         CFOP = IIf(Fatura_Proforma, 7101, 0)
         vgRetVal = CFOP
      Case 10  ' ← NOVO: Valor Unitário
         If Sequencia_do_Produto > 0 Then
            vgRetVal = InfoProdutos(..., "Valor Unitário")
         End If
   End Select
```

---

## OPÇÃO 3: Híbrida (usar ambas)

Usar `DEFAULTDASCOLUNAS` como principal, mas adicionar fallback em `IniApDaCol` se necessário.

**Problema**: Mais complexo sem necessidade.

---

## RECOMENDAÇÃO: OPÇÃO 2 ⭐

### Por quê?

1. **Menos código**: 3 linhas por grid vs ~15 linhas + variáveis
2. **Usa mecanismo nativo**: Sistema já faz isso para CST/CFOP
3. **Sem estado**: Não precisa rastrear "anterior"
4. **Mais limpo**: Declarativo, não procedural
5. **Mais seguro**: Menos lugares para dar erro

### Único teste necessário

Precisamos confirmar que `DEFAULTDASCOLUNAS` é chamado quando:
- ✅ Nova linha (ACAO_INCLUINDO) - CERTEZA que funciona
- ❓ **Troca de item na coluna 1** - PRECISA TESTAR

Se `DEFAULTDASCOLUNAS` NÃO for chamado na troca de item, aí sim precisamos da Opção 1.

---

## Próximos passos

1. **Implementar Opção 2** (DEFAULTDASCOLUNAS) no Grid(3) primeiro
2. **Testar** se funciona na troca de item
3. Se funcionar → replicar para Grid(0) e Grid(1) ✅
4. Se NÃO funcionar → implementar Opção 1 ⚠️

**Começamos testando a Opção 2?**
