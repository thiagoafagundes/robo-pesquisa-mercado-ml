# Robô de inteligência de mercado para o Mercado Livre

Ferramenta em Python que consulta a API oficial do Mercado Livre, encontra os produtos mais vendidos de um nicho e indica quais valem a pena revender, com nota, veredito e calculadora de margem.

Projeto pessoal de **Thiago de Almeida Fagundes**, graduando em Engenharia Elétrica (IFNMG, dez/2026). Portfólio completo em [`docs/Portfolio_IA_Thiago_de_Almeida_Fagundes.pdf`](docs/Portfolio_IA_Thiago_de_Almeida_Fagundes.pdf).

---

## O problema

Quem começa a vender online costuma escolher produto por intuição. A pergunta que eu queria responder com dados:

> **Quais produtos vendem muito, têm pouca concorrência e ainda deixam margem de lucro depois das taxas?**

## Como usei IA

Usei IA generativa (Claude) como **copiloto de programação**:

- **Eu** defini o problema, as regras de negócio e os critérios de decisão.
- **A IA** escreveu a maior parte do código a partir dessas instruções.
- **Eu** executei, testei com dados reais, identifiquei os erros e conduzi as correções.

## Como funciona

| Etapa | O que faz |
|---|---|
| 1. Login | OAuth 2.0 com PKCE na API do Mercado Livre |
| 2. Categorias | Seleciona as subcategorias do nicho por palavras-chave, com lista de exclusão |
| 3. Coleta | Top 20 mais vendidos de cada categoria |
| 4. Enriquece | Preço, concorrência no catálogo e buscas em alta |
| 5. Avalia | Nota de 0 a 100 e veredito (VENDER, ANALISAR, EVITAR) com o motivo |
| 6. Entrega | Planilha Excel com calculadora de margem e histórico semanal |

**Nota** = 40% demanda (posição no ranking) + 30% pouca concorrência + 15% faixa de preço + 10% constância no topo + 5% ausência de loja oficial.

**Regras de negócio (definidas por mim):**
- Só recomenda **VENDER** com preço entre R$ 50 e R$ 300, até 20 concorrentes e sem loja oficial dominando.
- Produtos elétricos recebem alerta de **certificação do Inmetro**.
- **Calculadora de margem** com fórmulas no Excel: o usuário informa o custo do fornecedor e a planilha desconta taxa, frete e imposto, indicando se a margem passa de 20%.
- **Histórico semanal** para separar tendência (produto que fica semanas no topo) de moda passageira.

## Resultado com dados reais (primeira execução, 10/11/2026)

| Categorias | Produtos avaliados | VENDER | ANALISAR | EVITAR | Buscas em alta |
|---|---|---|---|---|---|
| 20 | 400 | 62 | 210 | 128 | 955 |

A planilha real não está neste repositório. O arquivo em [`exemplo/`](exemplo/) usa **dados simulados** só para mostrar o formato.

## O que os dados reais me ensinaram

A versão testada só com dados simulados parecia perfeita. A primeira execução real revelou cinco problemas:

| Problema encontrado | Correção |
|---|---|
| Erro de login: a aplicação exigia PKCE | Implementação do fluxo PKCE (S256) |
| "bateria" trouxe instrumentos musicais; "fio" trouxe telefones "sem fio" | Palavras revistas e lista de categorias excluídas |
| 108 de 400 anúncios num formato novo (MLBU) vieram sem preço e título | Tratamento específico para esse formato, com link para conferência |
| Concorrência subestimada: a API conta só vendedores do catálogo | Alerta na planilha e regra de conferir no site antes de comprar |
| O mesmo produto aparecia em mais de uma categoria | Remoção de duplicados, que também reduz as consultas à API |

**Aprendizado:** a IA acelera muito a construção, mas validar com dados reais e questionar os resultados continua sendo responsabilidade de quem usa.

## Como rodar

**No Google Colab (recomendado):** abra `Robo_Pesquisa_Mercado_ML.ipynb` no Colab e siga as células. O passo a passo completo, incluindo como criar a aplicação no Mercado Livre, está em [`docs/GUIA_DE_USO.md`](docs/GUIA_DE_USO.md).

**No computador:**
```bash
pip install -r requirements.txt
python pesquisa_mercado.py --simular   # exemplo com dados falsos, sem credenciais
python pesquisa_mercado.py             # dados reais (preencha APP_ID e SECRET_KEY no arquivo)
```

As credenciais nunca vão para o repositório: o token fica em `token_ml.json`, que está no `.gitignore`.

## Próximos passos

- Usar o histórico semanal para treinar um modelo de **machine learning** que estime quais produtos tendem a permanecer no topo.
- Usar um **modelo de linguagem** para ler perguntas e avaliações dos anúncios e encontrar reclamações recorrentes, que indicam espaço para um produto melhor.

## Tecnologias

Python · requests · openpyxl · API REST do Mercado Livre · OAuth 2.0 / PKCE · Google Colab e Google Drive · IA generativa
