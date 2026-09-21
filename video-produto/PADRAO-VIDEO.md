# Padrão de construção de vídeo de produto

Processo para produzir vídeo vertical de 30s para marketplace e redes sociais,
usando Higgsfield (imagem em movimento) + ElevenLabs (locução e trilha) + ffmpeg
(animações e mixagem).

Consolidado a partir de dois projetos: Kit 4 Piscas CG/Titan/Fan/Twister e
Kit Frente Farol CG 160 Titan. O histórico de como cada regra surgiu está em
[`HISTORICO.md`](HISTORICO.md).

---

## Fase 1 — Intake

O que você me manda antes de começar:

### Obrigatório

| Item | Detalhe |
|---|---|
| **Fotos do SEU produto** | 4 a 8 imagens, do produto que você realmente envia. Frente, traseira, detalhes, e o conjunto completo |
| **Destino** | Shopee, Mercado Livre, Instagram, TikTok — muda o enquadramento seguro e o tom |
| **Compatibilidade exata** | Modelos, cilindradas e faixa de anos, separando encaixe **direto** de **adaptável**. Erro aqui gera devolução |
| **Conteúdo da caixa** | Quantas peças, o que acompanha, e **o que não acompanha** |
| **Variações** | Cores ou versões, com o nome exato de cada uma |

### Do anúncio do concorrente

Serve para **estudar**, nunca para gerar imagem. O Higgsfield trava no molde da
peça que recebe de referência: gerar a partir da foto do concorrente mostra o
produto dele, não o seu.

| Item | Para que uso |
|---|---|
| Descrição copiada | Vocabulário do nicho, argumentos que eles usam, o que omitem |
| **Avaliações, principalmente as ruins** | A reclamação recorrente do concorrente vira o argumento central da sua peça |
| Fotos do anúncio | Referência de ângulo e composição, não de produto |
| Preço | Só se você quiser comparação explícita — e aí preciso confirmar que é verdade |

### Claims que preciso que você confirme

Não invento nada sobre o produto. Se quiser afirmar, me confirme:

- Material da carcaça e da lente
- Garantia e prazo
- Comparação de preço ou de quantidade com o concorrente
- Certificação, homologação, Inmetro
- Origem, nacional ou importado

O que não puder ser verificado olhando a foto, ou você confirma, ou fica de fora.

### Checagem da descrição do anúncio

Descrição de marketplace costuma vir com resto de copiar-colar. Antes de usar
como fonte, conferir:

- Avisos que não têm relação com o produto (no farol: aviso sobre retrovisores)
- Frases arriscadas em peça de reposição, como *"qualidade original"*
- Título escrito para indexação, que lista cores ou modelos que o produto não tem

---

## Fase 2 — Roteiro

Devolvo, antes de gastar qualquer crédito:

1. **Texto da narração** com marcação de tempo por bloco
2. **Descrição visual** de cada bloco
3. **Animações de texto**, em página de revisão navegável quando o design for novo

### Regras do roteiro

**30 segundos comporta 55 a 65 palavras.** Passou disso, a locução atropela.

**Estrutura de 5 blocos**, que casa com os cortes do vídeo:

| Bloco | Duração | Função |
|---|---|---|
| 1 | 0–4s | Gancho. Ataca a dor, não o produto |
| 2 | 4–10s | O que vem no kit, ou o diferencial principal |
| 3 | 10–18s | Detalhes técnicos e instalação |
| 4 | 18–25s | Variações (cores) ou compatibilidade |
| 5 | 25–30s | Compatibilidade e fechamento |

**Lista longa vai na tela, não na fala.** Modelos compatíveis, anos e nomes de
cor: o espectador lê mais rápido do que ouve.

**Lista do que vem no kit vai na fala.** É o argumento central — vale ouvir, e
cada item ganha um selo na tela no instante em que é dito.

**Uma cor como herói.** Quando o produto tem variações, gerar o vídeo inteiro em
uma cor só e mostrar as demais apenas no bloco de cores.

---

## Fase 3 — Geração de imagem (Higgsfield)

**Modelo:** `seedance_2_5`, modo `omni_reference`, 9:16, 1080p, `bitrate_mode: high`

**Blocos:** somar 30s. Custo aproximado de 9 créditos por segundo.

### O que sempre entra no prompt

- **Quantidade exata**, em caixa alta: `EXACTLY ONE assembly and no more`
- **Escala com âncora corporal**: dimensão em centímetros mais a relação com a mão.
  Peça pequena: *"held between thumb and two fingers"*. Peça grande: *"handled with TWO hands"*
- **Distância de câmera explícita**: `the ENTIRE hand must be visible in frame`
- **Proibições**: farol ou lente nunca acesos, sem texto, sem marca d'água
- **Mãos**: `exactly five fingers per hand`, `one single pair of hands`
- **Terço inferior livre** nos blocos que vão receber animação: `leave the lower third clean`

### Armadilhas conhecidas

**Molde misturado.** Se as referências tiverem duas versões do produto, o modelo
alterna entre elas. Separar as fotos por molde antes.

**Mãos.** Onde o modelo mais erra. Contar com uma regeração a mais.

**Emblema e logotipo pequenos.** Detalhe fino demais: o modelo joga para a quina,
inclina ou espelha. Descrever a posição com precisão ajuda — *"flat, perfectly
horizontal, just right of centre, immediately above the lens, NOT at the edge,
NOT tilted"* — mas não garante o lado. Aceitar quando estiver limpo e plausível.

**Limite de simultaneidade.** Enviar 5 blocos de uma vez devolveu *"Out of
credits"* com 4.000 créditos em saldo. Era limite de trabalhos simultâneos, não
de saldo. Enviar em ondas de 2 ou 3.

**Preset recomendado.** O servidor sugere um preset; passar `declined_preset_id`
evita a interrupção.

### QA obrigatório antes de mostrar

Folha de contato em grade a 1 fps. Conferir quantidade de peças, molde, mãos,
escala, enquadramento e posição de emblema. Para detalhe fino, montar
comparativo lado a lado com a foto de referência em vez de julgar no olho.

---

## Fase 4 — Locução (ElevenLabs)

**Modelo: `eleven_v3`. Nunca `eleven_multilingual_v2`.**

O `multilingual_v2` lê de forma neutra e ignora qualquer direção de
interpretação. Foi o que produziu a locução sem vida no primeiro projeto.

### As três alavancas do v3

| Alavanca | Exemplo |
|---|---|
| Marcação de interpretação | `[excited]`, `[enthusiastic]`, `[energetic]` |
| Ênfase por maiúscula | `Trocar só um NÃO resolve` |
| Pontuação de ritmo | Exclamação, travessão para batida |

**Voz padrão:** André — Voz Animada & Confiante, `zxk5LWW71eFBR9Cep81P`
(pt-brazilian, profissional).

**Custo:** 40 a 125 créditos por frase, 1 a 2 centavos de dólar.

### Homógrafos

Palavras com duas pronúncias saem erradas. **"Logo"** foi lido como advérbio de
tempo em vez de emblema. A solução é trocar a palavra — **"logotipo"** — em vez
de tentar forçar a pronúncia.

### Sempre medir antes de montar

Gerar as frases separadas, medir a duração e conferir se cabem na janela do
bloco. Até meio segundo atravessando o corte é aceitável: o áudio faz a ponte.
Mais que isso, encurtar o texto. Nunca acelerar a dicção.

Conferir que a última frase **termina antes do fim do vídeo** — antecipar a
entrada se necessário.

---

## Fase 5 — Animações de texto (ffmpeg, custo zero)

Queimadas com `drawtext`, `drawbox` e `overlay`. Não gastam crédito.

### Identidade visual

| Elemento | Valor |
|---|---|
| Vermelho | `#DC1A22` |
| Branco | `#FFFFFF` |
| Preto | `#121013` |
| Fonte | Barlow Condensed ExtraBold e Bold, caixa alta |

Fonte em `github.com/google/fonts/raw/main/ofl/barlowcondensed/`, embutida no render.

### Posicionamento

- Margem lateral: 76px em 1080 de largura (7%)
- Texto ancorado na base, última linha a 10–12% do rodapé
- Para Stories, subir para 20%: a faixa inferior é onde Reels e TikTok põem perfil e legenda

### O que o ffmpeg reproduz de verdade

Fade por `alpha` no `drawtext`, deslizar por expressão em `x`/`y`, entrada em
sequência com `enable=between(t,a,b)`, barra preenchendo com `drawbox` de largura
variável, tremor com `sin()` amortecido por `exp()`.

**Não tentar:** texto crescendo, girando ou deformando.

**`overlay` não aceita `alpha` animado** — o parâmetro é seletor de modo, não
valor. Para imagens sobrepostas, a entrada em sequência vem do `enable`.

### Regras de conteúdo na tela

**Acentos:** passar texto acentuado por `textfile=` em UTF-8. Inline no filtro,
o parser come os acentos.

**Cores de variação:** círculos com a cor e o nome, nunca moto gerada por IA.
Modelo de imagem erra grafismo de moto real, e o comprador conhece.

**Marcador aponta para linha.** Rótulo como `ADAPTÁVEL` precisa estar ao lado da
linha a que se refere. Solto na tela, parece valer para tudo.

**Aviso honesto em texto pequeno.** O que não acompanha o produto (presilhas, por
exemplo) entra no vídeo. Dizer antes da compra derruba reclamação depois.

### O que foi recusado e não deve voltar

- Contador `1 · 2 · 3 · 4` — mostrar o número direto
- Silhueta de moto desenhada ou animada — escrever só o modelo
- Adjetivo técnico vazio, como *"facetada"*

---

## Fase 6 — Trilha e mixagem

**Modelo:** `eleven_music_v2`, instrumental. **900 créditos, cerca de 15 centavos**
por faixa.

### O que pedir no prompt

- Estilo, BPM, tonalidade e instrumentação
- `groove starts immediately, no long intro`
- `midrange kept open for a voiceover`
- `energy stays completely constant, no build-up, no drop`
- `no vocals, fully instrumental`
- O que **não** quer: distorção pesada, riser, 808, sirene

### Trilha a partir de uma referência

Quando existir uma música de referência:

1. Extrair o áudio e medir com `librosa`: BPM, tônica provável, distribuição
   espectral por faixa, centroide, e variação de energia ao longo do tempo
2. Escrever o prompt com essas medidas — andamento, tom, onde o mix pesa,
   se a energia é constante
3. Gerar duas variações com instrumentação diferente

**Estilo não tem direito autoral; melodia e composição têm.** O prompt descreve
características, nunca a melodia. Clonar a linha melódica faz o anúncio ser mutado
ou derrubado por detecção automática.

**O áudio de referência nunca vai para repositório público.** Guardar localmente.

### Balanço medido, não no ouvido

| Alvo | Valor |
|---|---|
| Música abaixo da fala | 10 a 12 dB |
| Nível nos trechos falados | −20 a −15 dB |
| Nível nos intervalos | −31 a −25 dB |

Método: `volume=-14dB` na música, mais `sidechaincompress` com a locução como
chave, para a música recuar sozinha quando o locutor fala. Entrada de 0,4s e
saída de 1,8s. Filtro em
[`projetos/02-frente-farol-cg160/montagem/mix-audio.txt`](projetos/02-frente-farol-cg160/montagem/mix-audio.txt).

**Entrega:** `loudnorm=I=-16:TP=-1.5:LRA=11`.

---

## Especificação de entrega

| Item | Valor |
|---|---|
| Resolução | 1080×1920 (9:16) |
| Frame rate | 24 fps |
| Duração | 30s |
| Vídeo | H.264, CRF 18, preset slow, profile high, level 4.1, `+faststart` |
| Áudio | AAC 192k, 48 kHz, −16 LUFS, pico −1,5 dBTP |
| Tamanho típico | 14 a 19 MB |

---

## Custo por projeto

| Etapa | Custo | Refazer |
|---|---|---|
| Intake e roteiro | zero | livre |
| Blocos de vídeo | 270 a 370 créditos Higgsfield | caro — acertar de primeira |
| Locução | 5 a 10 centavos USD | barato |
| Animações e mixagem | zero | livre |
| Trilha | 15 centavos USD por faixa | barato |

**A geração de vídeo é a única etapa cara.** Por isso roteiro e referências são
aprovados antes, e o QA acontece antes de mostrar qualquer coisa.
