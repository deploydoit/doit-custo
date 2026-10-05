# Como funciona a atualização de custos da Luminata

Guia rápido para quem vai usar o app **Atualização de Custos** (doit-custo). O objetivo do app é pegar a **tabela de preços de um fornecedor** e transformar nos **custos** dos produtos correspondentes no **DOit**.

---

## 1. A conta do custo (o coração da lógica)

Para cada produto, o cálculo é feito em duas etapas:

**Passo 1 — Custo Líquido**
```
Custo Líquido = Preço do fornecedor × 1,10
```
Ou seja, aplicamos **10%** sobre o preço que veio na planilha do fornecedor.

**Passo 2 — Custo Bruto (o "CUSTO" que vai pro DOit)**
```
Custo Bruto = Custo Líquido × (1 + IPI%)
```
Aqui aplicamos o **IPI** por cima do custo líquido.

> Ambos os valores são arredondados para **2 casas decimais**.

### Exemplo prático
Fornecedor manda um produto por **R$ 100,00**, com IPI de **10%**:

| Etapa | Conta | Resultado |
|---|---|---|
| Custo Líquido | 100,00 × 1,10 | **110,00** |
| Custo Bruto | 110,00 × 1,10 | **121,00** |

No arquivo final, esse produto sai com **CUSTO LÍQUIDO = 110,00** e **CUSTO = 121,00**.

---

## 2. O IPI: fixo ou por produto

O IPI pode entrar de duas formas:

- **IPI fixo:** você digita um único valor (%) e ele vale para todos os itens daquele fornecedor.
- **IPI por produto:** quando a planilha do fornecedor já traz o IPI item a item (ex: **Stella**, coluna `% IPI`). O app usa o IPI de cada linha. Se algum item vier **sem IPI**, ele usa o IPI fixo como reserva.

> Regra de bolso: se a planilha tem uma coluna de IPI por item, use "IPI por produto". Senão, digite o IPI fixo.

### Quando o preço já é o custo líquido

Alguns fornecedores já mandam o **custo líquido** direto na tabela, então o Passo 1 (× 1,10) **não** deve ser aplicado. Para esses casos existe a opção **"Preço já é o custo líquido"**. Com ela ligada a conta vira:

```
Custo Líquido = Preço do fornecedor        (sem o × 1,10)
Custo Bruto   = Custo Líquido × (1 + IPI%)
```

> A **Accord** usa um caso ainda mais específico (dois acabamentos e **sem IPI**). Veja o item 7.

---

## 3. Como o app sabe qual produto é qual (cruzamento)

O app cruza o **código do fornecedor** com a **referência (`# Referência`) do produto no DOit**.

Pontos importantes:

- **Sufixo do DOit:** muitas referências no DOit vêm como `CÓDIGO | 394` (código + um número interno). O app compara **só a parte antes do `|`**, então `SD1000BR` casa com `SD1000BR | 394`.
- Cada fornecedor tem particularidades (código com acabamento, código montado a partir de duas colunas, etc.). Por isso existe o seletor **"Qual é o fornecedor?"**, que já pré-configura tudo (aba, cabeçalho, colunas, opções e o fabricante no DOit).

---

## 4. Os três grupos do resultado

Depois de processar, o app separa os produtos em:

- ✅ **Atualizados** — casaram com o DOit e tiveram o custo recalculado.
- 🆕 **Precisam ser criados** — estão na planilha do fornecedor, mas ainda **não existem** no DOit.
- ⚠️ **Não atualizados** — existem no DOit para aquele fabricante, mas **não vieram** na planilha do fornecedor.

---

## 5. O arquivo final (Importação DOit)

A aba **Modelo Custo** é o que se importa no DOit. Colunas principais:

| Coluna | O que é |
|---|---|
| **SKU** | Código do produto no DOit (texto — preserva zeros à esquerda) |
| **FORNECEDOR** | Id do fabricante no DOit |
| **CONDIÇÃO** | Sempre `DDP` |
| **CUSTO** | Custo Bruto (líquido + IPI) — número |
| **MOEDA** | Sempre `BRL` |
| **CUSTO LÍQUIDO** | Preço × 1,10 — número |
| **MODIFICADO EM** | Data do processamento |

> CUSTO e CUSTO LÍQUIDO saem como **número** e a data como **data** — não precisa converter nada no Excel.

---

## 6. Passo a passo resumido no app

1. Faça o **upload** da planilha do fornecedor.
2. Escolha o **fornecedor** na lista e clique em **"Aplicar configuração"** (preenche aba, colunas e opções).
3. Confira **Coluna do CÓDIGO**, **Coluna do PREÇO** e o **Fabricante no DOit**.
4. Defina o **IPI** (fixo ou por produto).
5. Clique em **"Processar atualização"**.
6. Baixe a **Importação DOit** e confira os grupos (atualizados / criar / não atualizados).
7. Preencha o **Registro** e cole na planilha de controle (Google Sheets).

---

## 7. Particularidades da Accord

A Accord (fabricante **L Martinkoski E Cia Ltda Me (Accord)**, ID `7623` no DOit) é um caso à parte. A receita pronta já resolve tudo, mas vale entender a lógica:

- **Várias abas:** a planilha tem 6 abas de produto (**Pendente, Plafon, Arandela, Abajur, Coluna, Mobiliário**) e 3 informativas (Opções de Acabamento, Informações LED, Política de Compra e Venda). A receita já **seleciona apenas as abas de produto** — as informativas são ignoradas automaticamente (não precisa tirar nada à mão).
- **Cabeçalho na linha 1** e colunas pela posição: **Código = coluna C (Referência)**, **Lâmina Natural = coluna E** e **Lâmina Tingida = coluna F**.
- **Cruzamento direto:** a referência da Accord (ex: `116`, `202LED`, `4234`) bate direto com a `# Referência` do DOit.

### Dois acabamentos: Natural × Tingida

Cada referência tem **dois preços** (Lâmina Natural e Lâmina Tingida) e, no DOit, vira **vários SKUs** — um por acabamento. Os dois são produtos diferentes, com descrições e custos diferentes, então o app diferencia por SKU:

- SKU com **"NATURAL"** no nome → recebe o preço da **Lâmina Natural**.
- SKU de **cor/tingido** (ex: 33-Bronze, 40-Azul, 42-Curupixá) → recebe o preço da **Lâmina Tingida**.

Se a **Lâmina Tingida** vier **vazia** para uma referência, os SKUs de cor dela **ficam sem atualizar** (vão para "não atualizados").

### Sem IPI — os 10% são margem

A Accord **não tem IPI**. Os 10% são **margem** aplicada sobre o preço do acabamento:

```
Custo Líquido = preço do acabamento (Natural ou Tingida)
Custo         = Custo Líquido × 1,10
```

> **Atenção à aba Coluna:** ela traz valores de IPI (7,5% e 10%) no topo. Esses valores devem ser **ignorados** — o padrão da Accord é sem IPI.

> Como a planilha nasce de cálculos de markup, muitos preços têm **3 casas decimais** (ex: `1394,275`). O app trata esses valores como número e **não** os confunde com separador de milhar.

---

### Resumo de uma frase
> **Custo = (Preço do fornecedor × 1,10) × (1 + IPI).** O resto do app é só garantir que cada produto da planilha encontre o produto certo no DOit.
>
> *(Exceções: fornecedores cujo preço já é o custo líquido pulam o primeiro × 1,10; a Accord não tem IPI e escolhe o preço por acabamento — Natural ou Tingida — de cada SKU.)*
