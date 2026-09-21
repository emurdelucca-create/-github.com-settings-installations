# Vídeo de produto

Pipeline de produção de vídeos verticais de 30s para marketplace (Shopee, Mercado
Livre) e redes sociais, para peças de moto.

## Por onde começar

| Arquivo | Para que serve |
|---|---|
| [`PADRAO-VIDEO.md`](PADRAO-VIDEO.md) | **O processo.** Intake, roteiro, geração, locução, animações, trilha, entrega |
| [`HISTORICO.md`](HISTORICO.md) | Como o processo foi construído: cada problema e a decisão que ele gerou |
| [`projetos/`](projetos/) | Os produtos já feitos, com roteiro, referências e vídeo final |
| [`ferramentas/`](ferramentas/) | Página de revisão de animações |

## Projetos

| # | Produto | Vídeo final |
|---|---|---|
| 01 | [Kit 4 Piscas CG/Titan/Fan/Twister](projetos/01-kit-4-piscas/) | `kit4piscas_v7_30s_9x16_final.mp4` |
| 02 | [Kit Frente Farol CG 160 Titan](projetos/02-frente-farol-cg160/) | `farol_cg160_30s_9x16_COMPLETO.mp4` |

## Stack

| Etapa | Ferramenta |
|---|---|
| Blocos de vídeo | Higgsfield — `seedance_2_5`, modo `omni_reference` |
| Locução | ElevenLabs — `eleven_v3`, voz André |
| Trilha | ElevenLabs — `eleven_music_v2` |
| Animações e mixagem | ffmpeg |

## Como fazer um produto novo

Mandar para o Claude:

1. Fotos **do seu produto** — frente, traseira, detalhes, conjunto
2. Destino do vídeo
3. Compatibilidade exata, separando encaixe direto de adaptável
4. O que vem e o que **não** vem na caixa
5. Variações de cor, com os nomes
6. Descrição e avaliações do concorrente, se quiser — servem para estudo

Ele devolve o roteiro antes de gastar crédito. Detalhes em
[`PADRAO-VIDEO.md`](PADRAO-VIDEO.md), Fase 1.

## Reproduzir a montagem de um vídeo

Os filtros de montagem estão em `projetos/<produto>/montagem/`. Com os blocos de
vídeo, as frases de locução e os arquivos de texto em mãos:

```bash
ffmpeg -i base.mp4 \
  -loop 1 -i c1.png -loop 1 -i c2.png -loop 1 -i c3.png \
  -loop 1 -i c4.png -loop 1 -i c5.png -loop 1 -i c6.png \
  -i n1.mp3 -i n2.mp3 -i n3.mp3 -i n4.mp3 -i n5.mp3 \
  -filter_complex_script filtro-video.txt \
  -map "[vout]" -map "[aout]" -t 30.209 \
  -c:v libx264 -preset slow -crf 18 -pix_fmt yuv420p \
  -profile:v high -level 4.1 -r 24 \
  -c:a aac -b:a 192k -ar 48000 -movflags +faststart saida.mp4
```

E para somar a trilha, copiando o vídeo sem recodificar:

```bash
ffmpeg -i saida.mp4 -i trilha.mp3 \
  -filter_complex_script mix-audio.txt \
  -map 0:v -c:v copy -map "[aout]" -c:a aac -b:a 192k -ar 48000 \
  -movflags +faststart final.mp4
```
