# Histórico da produção

Como o processo foi construído, na ordem em que as coisas aconteceram. Cada
problema encontrado aqui virou uma regra em [`PADRAO-VIDEO.md`](PADRAO-VIDEO.md).

Produção feita entre 18 e 19 de setembro de 2026, com Claude operando Higgsfield,
ElevenLabs e ffmpeg.

---

## Projeto 1 — Kit 4 Piscas CG/Titan/Fan/Twister

### Versão 1: estúdio preto e locução apertada

Primeiro briefing pedia vídeo de 30s com fundo preto de estúdio, sem mãos, e
locução em português.

**Problema de referência.** As fotos enviadas tinham **dois moldes diferentes**
do pisca: lente cônica e lente em cunha. O primeiro render alternou entre os dois
e também enfileirou oito piscas em vez de quatro. Refeito travando só nas
referências da cunha e escrevendo a quantidade em caixa alta.

**Problema de tempo.** A copy, lida em ritmo natural, dava 41,9s para um vídeo de
30s. A dicção foi acelerada para caber, e perdeu o respiro.

> Regra que nasceu daqui: 30s comporta 55 a 65 palavras.

**Sem trilha.** O conector do Higgsfield só gerava fala; música e efeitos eram
restritos a outro pipeline. A v1 saiu com locução sobre imagem sem música.

### Versão 2: bancada e mãos

Retorno: os takes na moto ficaram ótimos, mas o fundo preto ficou artificial.
Pedido: pessoa manuseando a peça, câmera mais afastada, e vídeo sem som.

- Cenário escolhido: bancada de oficina, só mãos e antebraços
- **Câmera colada na peça** nos primeiros renders — os macros extremos foram
  descartados e o prompt passou a exigir a mão inteira visível no quadro
- **Escala** calibrada em duas rodadas; a versão final deixa o pisca um pouco
  maior que o real em relação à mão, escolha para leitura em e-commerce
- Takes da moto aprovados foram **reaproveitados sem regerar**, cortando só os
  3s finais em fundo preto

### Versão 3 a 5: narração e animações

**Página de revisão.** Antes de queimar qualquer animação, montei uma página com
os cinco momentos animados sobre frames reais do vídeo. O texto desceu para a
base da tela a pedido.

**Primeira locução rejeitada.** O TTS do Higgsfield soou sintético. Retorno do
usuário: a falta de uma narração em português de verdade deixava o vídeo ruim.

**Animações rejeitadas:**

- Contador `1 · 2 · 3 · 4` — trocado por **KIT COM 4** direto
- Silhueta de moto animada — trocada pelos nomes dos modelos em texto
- *"Lente âmbar facetada"* — tirado o "facetada"
- *"Terminais bala"* — trocado por **"fiação completa"**

**Acentos sumidos.** O primeiro render saiu com `NAO` e `AMBAR`: os acentos
tinham sido removidos para escapar do parser do ffmpeg. Resolvido passando o
texto por arquivo UTF-8.

### CapCut e a descoberta do ElevenLabs

Pedido para conectar o CapCut. **Não existe conector do CapCut**, e ele não tem
API pública de narração. Mas o ElevenLabs já estava conectado na sessão — e é
o motor que muitos editores usam por baixo.

Seis vozes brasileiras testadas lado a lado com o mesmo texto. Escolhida:
**André — Voz Animada & Confiante**.

### Versão 6: a voz "morta"

Retorno: a voz estava sem vida. Diagnóstico: o problema era o **modelo**, não a
voz. O `eleven_multilingual_v2` ignora direção de interpretação. Troca para o
`eleven_v3` com marcações `[excited]`, ênfase em maiúscula e pontuação de ritmo.

> Regra: sempre `eleven_v3`.

### Versão 7: trilha

Três trilhas geradas no ElevenLabs — rock de garagem, leito comercial e hip-hop
contido — e cada uma mixada no vídeo para escolha no contexto real. Escolhida a
de rock de garagem.

Pedido de cuidado com o volume. A primeira mixagem deixava a música 8 dB abaixo
da voz; desceu para 11–12 dB, com abaixamento automático por sidechain.

---

## Criação do padrão

Com o primeiro projeto fechado, o processo foi consolidado em `PADRAO-VIDEO.md`
para ser repetido em outros produtos.

---

## Projeto 2 — Kit Frente Farol CG 160 Titan

### Intake

Material: fotos do produto, fotos das motos em cada cor, descrição do anúncio do
concorrente.

**Imagens do concorrente reposicionadas.** O plano inicial era gerar o vídeo a
partir delas. Ajustado: servem para estudar o anúncio, e o vídeo é gerado a
partir das fotos do produto real.

**Descrição com problemas herdados:**

- Aviso legal sobre **retrovisores**, sem relação com o produto
- *"Produto com qualidade original"* numa peça de reposição — retirado da fala
- Título listando cores que não batiam com as reais — era só para indexação

**Confirmações do usuário:**

- Seis cores: Preto Brilhante, Azul Twister Metálico, Amarelo Caju, Azul Caraiva,
  Branco Andes, Prata Force
- Logotipo HONDA prateado
- "Bananinha" = abas laterais coloridas do farol, que é o que muda de cor
- O azul das fotos é o Twister

**Decisão de cor.** Vídeo inteiro na cor Preto Brilhante; as seis cores aparecem
só no bloco de cores, como círculos com nome — não como motos geradas por IA.

### Geração

**Limite de simultaneidade.** Cinco blocos enviados de uma vez: três falharam com
*"Out of credits"*, apesar de 4.023 créditos em saldo. Reenviados em ondas
menores, entraram todos.

**QA dos blocos.** O bloco 3 entregou exatamente o pedido: a peça virando para
revelar a aranha traseira, depois os parafusos na bancada.

**Emblema HONDA no lugar errado.** Apontado pelo usuário. Comparativo lado a lado
com a referência confirmou: no bloco 1 estava na quina e inclinado; no bloco 2,
do lado oposto. Regerados com a posição descrita em detalhe — saíram deitados e
no visor, mas espelhados. **Aceito pelo usuário** sem nova tentativa.

### Locução

**"Logo" lido como advérbio.** Apontado pelo usuário: a fala soava como "logo
mais". Trocado por **"logotipo Honda"**.

### Montagem

- `overlay` rejeitou `alpha` animado — a entrada dos círculos passou a ser pelo `enable`
- Rótulo `ADAPTÁVEL` ficou solto na tela, dando a entender que valia também para
  a Titan 2022, que é encaixe direto. Corrigido com marcador por linha

### Trilha a partir de referência

Pedido: música sem direito autoral baseada num vídeo de referência.

O áudio foi medido, não copiado: **123 BPM**, tônica provável em **F#**, mix
concentrado nos médios com pouco sub-grave, e **energia constante** do início ao
fim dos 35s. Duas trilhas geradas a partir desse perfil. Escolhida a primeira,
de synth plucado com guitarra limpa.

> Estilo não é protegido por direito autoral; melodia é. O vídeo de referência
> não foi incluído neste repositório.

Uma comparação numérica entre o perfil da referência e o das trilhas geradas
chegou a ser iniciada, mas **não foi concluída**: o ambiente de processamento
reiniciou e perdeu o áudio extraído. Como a trilha já tinha sido aprovada de
ouvido, não foi refeita.

---

## Ferramentas usadas

| Ferramenta | Papel |
|---|---|
| Higgsfield — `seedance_2_5` | Geração dos blocos de vídeo |
| ElevenLabs — `eleven_v3` | Locução |
| ElevenLabs — `eleven_music_v2` | Trilhas |
| ffmpeg | Concatenação, animações de texto, mixagem |
| librosa | Análise da música de referência |
| rsvg-convert | Círculos de cor em PNG |
