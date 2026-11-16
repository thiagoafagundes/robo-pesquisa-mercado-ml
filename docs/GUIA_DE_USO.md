# Robô de Pesquisa de Mercado: passo a passo (versão 2)

## O que mudou na versão 2
- **Foco no seu nicho:** ele escolhe sozinho as subcategorias de elétrico, ferramentas, iluminação, energia, carregadores e afins.
- **Regras que combinamos:** só dá **VENDER** com preço entre R$ 50 e R$ 300, até 20 concorrentes e sem loja oficial dominando.
- **Inmetro:** marca "Verificar" nos produtos que costumam exigir certificação, como tomadas, disjuntores, carregadores e extensões.
- **Calculadora de margem:** você escreve só o custo do fornecedor (coluna amarela). A planilha calcula taxa do ML, frete, imposto, lucro e margem, e diz se **passa nos 20%**.
- **Histórico semanal:** a coluna "Vezes no topo" mostra em quantas semanas o produto apareceu. 3 ou mais é tendência.

## 1. Crie seu "aplicativo" no Mercado Livre (uma vez só, uns 10 min)
1. Entre em **developers.mercadolivre.com.br** com a sua conta do Mercado Livre.
2. Vá em **Minhas aplicações → Criar aplicação**. Os nomes dos botões podem mudar um pouco.
3. Preencha os campos:
   - **Nome / nome curto / descrição:** qualquer um. Ex.: "Pesquisa Thiago".
   - **URI de redirect:** `https://www.google.com/`
   - **PKCE:** tanto faz, o robô funciona ligado ou desligado.
   - **Permissões (escopos):** marque **leitura** e, se aparecer, **offline_access**. É ela que deixa o robô renovar o login sozinho.
   - **Não precisa** marcar tópicos nem notificações.
4. Salve e copie o **App ID** (Client ID) e a **Secret Key** (Client Secret).
   - **A Secret Key é como uma senha:** não mande para ninguém.

## 2. Rode no Google Colab (sem instalar nada)
1. Abra **colab.research.google.com**, entre com sua conta Google e vá em **Arquivo → Fazer upload de notebook**.
2. Escolha `Robo_Pesquisa_Mercado_ML.ipynb`.
3. Clique no ▶ de cada célula, **de cima para baixo**:
   - **Passo 1:** conecta seu Google Drive. Login e histórico ficam na pasta **RoboML**.
   - **Passo 2:** instala o robô.
   - **Passo 3:** cole o App ID e a Secret Key.
   - **Passo 4:** abra o link, clique em **Permitir** e cole o endereço inteiro da página do Google (o que tem `?code=TG-...`).
   - **Passo 5:** a pesquisa roda (3 a 10 minutos).
   - **Passo 6:** baixa a planilha.
4. **Nas semanas seguintes:** rode os Passos 1, 2, 3, 5 e 6. Só refaça o Passo 4 se der erro de login.

## 3. Como usar a planilha
1. Aba **Ranking**: os produtos **VENDER** estão no topo.
2. Para cada candidato, ache um fornecedor no Brasil com nota fiscal e escreva o **custo por unidade** na coluna amarela.
3. Olhe **"Passa nos 20%?"**. Se for NÃO, descarte ou negocie preço com o fornecedor.
4. Se a coluna **Inmetro** disser "Verificar", só compre produto com selo e certificado do fornecedor.
5. Prefira o que tem **"Vezes no topo" ≥ 3**.
6. Teste com 5 a 10 unidades de 2 ou 3 produtos por 30 a 45 dias. Depois aumente o que girou com lucro.

**Confira as taxas:** os valores de taxa do ML, custo fixo abaixo de R$ 79, frete e imposto ficam na aba **Parâmetros** e são estimativas. Ajuste-os com os valores atuais do Mercado Livre e da sua empresa (MEI = imposto 0). Todas as contas se atualizam.

## 4. Rodar no computador (opcional)
1. Instale o Python (python.org).
2. No terminal, rode: `pip install requests openpyxl`
3. Preencha `APP_ID` e `SECRET_KEY` no topo do `pesquisa_mercado.py`.
4. Rode `python pesquisa_mercado.py`. Depois da 1ª autorização, ele lembra o login.
5. Guarde o arquivo `historico_mais_vendidos.csv` na mesma pasta: é a memória das semanas.

**Ver a planilha de exemplo sem credenciais:** `python pesquisa_mercado.py --simular` (os dados são **falsos**).

## 5. Erros comuns
| Mensagem | O que fazer |
|---|---|
| `redirect_uri` / `invalid_grant` | A URI precisa ser idêntica à do aplicativo, ou o code venceu. Gere um link novo (Passo 4). |
| "Token inválido ou vencido" / "Sem token válido" | Refaça o Passo 4. |
| "sem ranking de mais vendidos nesta categoria" | É normal. Algumas categorias não têm ranking. |
| "Nenhuma subcategoria bateu com o NICHO" | Revise as palavras da lista `NICHO` no Passo 3. |
