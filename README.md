# PLANO DE CORTE SLITTER


## 1. O Que o Script Faz — Visão Geral

O script resolve um problema clássico da indústria de tubos e perfis: **como aproveitar ao máximo a largura de uma bobina de aço**, combinando cortes de diferentes perfis (matrizes) de forma que a perda de material fique dentro de uma faixa aceitável.

O usuário informa um perfil principal chamado **âncora** — obrigatoriamente produzido — e o script descobre quais outros perfis (**complementares**) preenchem o espaço restante da bobina dentro dos limites de perda.

### Analogia

Pense em uma prateleira de 1200 cm. Você precisa colocar caixas de tamanhos variados. O objetivo é que a soma das caixas deixe uma folga entre **0,67% e 1,70%** do espaço total — nem pouco, nem muito.

---

## 1.1 Conceitos Fundamentais

| Termo                | Significado                                                               |
| -------------------- | ------------------------------------------------------------------------- |
| Matriz               | Perfil de tubo ou barra (ex: 50,80-2", 100X50). Define o desenvolvimento. |
| Desenvolvimento (mm) | Largura da tira de aço necessária para fabricar o perfil.                 |
| Largura da bobina    | Largura fixa do aço: 1000, 1200 ou 1500 mm.                               |
| Âncora               | Perfil obrigatório, com N ≥ 1 corte.                                      |
| Complementar         | Perfis que ocupam o espaço restante da bobina.                            |
| Combinação           | Âncora(N) + Comp1(N1) + Comp2(N2).                                        |
| Total de cortes      | Soma total dos cortes da combinação.                                      |
| Perda (%)            | (Largura − Soma dos cortes) / Largura × 100.                              |

---

## 2. Parâmetros de Negócio

```python
LARGURAS_BOBINA = [1200, 1000, 1500]
PERDA_MIN_PCT = 0.67
PERDA_MAX_PCT = 1.70
MAX_COMP_NA_COMBO = 2
PESO_MEDIO_BOB_PAD = 12_000
QTD_BOBINAS_PAD = 1
```

| Constante                       | O que controla     | Exemplo de ajuste |
| ------------------------------- | ------------------ | ----------------- |
| `LARGURAS_BOBINA`               | Ordem das larguras | Adicionar `1800`  |
| `PERDA_MIN_PCT / PERDA_MAX_PCT` | Janela de perda    | `0.50 – 2.00`     |
| `MAX_COMP_NA_COMBO`             | Complementares     | `3` (mais lento)  |
| `PESO_MEDIO_BOB_PAD`            | Peso padrão        | `15000`           |
| `QTD_BOBINAS_PAD`               | Qtd. padrão        | `1`               |

---

## 3. Fluxo Completo de Execução

| Etapa | O que acontece              |
| ----- | --------------------------- |
| 1     | Carga do Excel e validação  |
| 2     | Escolha da espessura        |
| 3     | Escolha do tipo de material |
| 4     | Escolha da matriz âncora    |
| 5     | Limite de cortes            |
| 6     | Quantidade de bobinas       |
| 7     | Peso do lote                |
| 8     | Busca de combinações        |
| 9     | Motor combinatorial         |
| 10    | Exibição no terminal        |
| 11    | Exportação para Excel       |

---

## 4. Interface de Usuário — Menu

### Passos 1, 2 e 3

Filtro progressivo por **Espessura → Tipo → Âncora**.

### Passo 4 — Limite de Cortes

Restrição física da máquina.

| Combinação            | Total de cortes | Resultado |
| --------------------- | --------------- | --------- |
| Âncora(x3) + Comp(x2) | 5               | ✓         |
| Âncora(x2) + Comp(x4) | 6               | ✗         |
| Âncora(x1)            | 1               | ✓         |

> Se deixado em branco, não há limite.

### Passo 5 — Quantidade de Bobinas

Define o número de bobinas do lote.

### Passo 6 — Peso das Bobinas

Peso total do lote em kg.

**Fórmula do peso médio:**

```
Peso médio = Peso informado / Quantidade de bobinas
```

---

## 5. Lógica de Busca

### 5.1 Seleção da Largura

```python
for largura in [1200, 1000, 1500]:
    buscar_combinacoes()
    if encontrou:
        break
```

### 5.2 Validação

* 0,67% ≤ Perda ≤ 1,70%
* Total de cortes ≤ limite (se houver)

### 5.3 Motor Combinatorial

```python
para cada N_âncora:
    testar âncora
    testar âncora + complementares
```

---

## 6. Cálculo de KG

**Etapa 1:**

```
peso_medio = peso_total / qtd_bobinas
```

**Etapa 2:**

```
KG_i = (peso_medio / largura_bobina) * (N_cortes * desenvolvimento * qtd_bobinas)
```

**Etapa 3:**

```
KG_total = soma(KG_i)
```

---

## 7. Saídas

### Terminal

Resumo + combinações válidas ordenadas por perda.

### Excel

#### Aba Combinações

| Campo        | Descrição        |
| ------------ | ---------------- |
| #            | ID               |
| Combinação   | Descrição        |
| N Âncora     | Cortes da âncora |
| Total Cortes | Soma             |
| Perda (%)    | Perda            |
| Qtd KG       | Quilos           |

#### Aba Detalhes

| Campo           | Descrição           |
| --------------- | ------------------- |
| # Combo         | ID                  |
| Papel           | Âncora/Complementar |
| Matriz          | Perfil              |
| Desenvolvimento | mm                  |
| N° Cortes       | Quantidade          |
| Subtotal        | mm                  |
| Qtd KG          | kg                  |

---

## 8. Personalização

### Ajustes rápidos

| O que mudar     | Onde               | Como            |
| --------------- | ------------------ | --------------- |
| Janela de perda | PERDA_MIN/MAX      | Alterar valores |
| Larguras        | LARGURAS_BOBINA    | Reordenar       |
| Complementares  | MAX_COMP_NA_COMBO  | Aumentar        |
| Peso padrão     | PESO_MEDIO_BOB_PAD | Ajustar         |
| Qtd bobinas     | QTD_BOBINAS_PAD    | Ajustar         |

### Avisos

* MAX_COMP_NA_COMBO = 3 impacta performance
* Arquivo de entrada fixo
* Limite muito restritivo pode gerar vazio
* Peso sempre do lote total

---

## 9. Glossário

| Termo            | Significado        |
| ---------------- | ------------------ |
| DataFrame        | Tabela em memória  |
| combinations     | Subconjuntos       |
| product          | Produto cartesiano |
| int | None       | Valor ou ausência  |
| groupby().mean() | Média por grupo    |
| dropna()         | Remove nulos       |
| astype().strip() | Limpeza de texto   |
| pd.to_numeric    | Conversão numérica |

---

