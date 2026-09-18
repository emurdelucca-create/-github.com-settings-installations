# Vídeo de produto — Kit 4 Piscas CG/Titan/Fan

Vídeo vertical para marketplace e redes sociais do kit de 4 piscas (seta) para
Honda CG/Titan/Fan 125/150, modelo de lente em cunha com topo preto.

## Entregáveis

| Arquivo | Codec | Tamanho | Áudio | Uso |
|---|---|---|---|---|
| `kit4piscas_v2_30s_9x16_mudo.mp4` | H.264 | 12,9 MB | nenhum | **versão atual aprovada** |
| `kit4piscas_30s_9x16_h264.mp4` | H.264 | 11,1 MB | locução pt-BR | versão 1, substituída |
| `kit4piscas_30s_9x16_hevc.mp4` | H.265 | 43,2 MB | locução pt-BR | master da versão 1 |

Todos em 1080×1920 (9:16), 24 fps, 30 s.

## Estrutura da versão 2

Corte seco entre blocos, sem áudio.

| Trecho | Conteúdo |
|---|---|
| 0–10 s | Bancada de oficina; mão pega o pisca, gira mostrando a lente âmbar e o perfil |
| 10–18 s | Traseira e haste, conjunto porca/parafuso, chicote esticado, terminais bala |
| 18–26 s | Instalado na CG verde: close no farol, travelling lateral, frontal 3/4 |
| 26–30 s | As quatro unidades nas duas mãos |

## Decisões de produção

**Fundo preto removido.** A versão 1 usava fundo preto de estúdio nos blocos de
produto, que ficou artificial. A versão 2 substituiu por manuseio em bancada de
oficina, com mãos e antebraços apenas — sem rosto, para reduzir artefato.

**Enquadramento.** Câmera mantida a distância média, com a mão inteira sempre no
quadro. Macros extremos foram descartados. O pisca aparece um pouco maior do que
a proporção real em relação à mão (peça real ~10 cm, mão ~19 cm), escolha
deliberada para leitura em e-commerce.

**Takes da moto preservados.** O bloco na CG foi aprovado na versão 1 e
reaproveitado sem regerar; apenas os ~3 s finais em fundo preto foram cortados.

**Sem áudio.** A versão 1 tinha locução pt-BR normalizada em −16 LUFS. A versão 2
foi entregue muda por decisão de briefing.

## Atenção ao molde

As imagens em `referencias/` contêm **dois moldes diferentes** do produto:

- `br-11134207-820ma-*` e `br-11134207-820mb-*` — lente **cônica**
- as demais — lente em **cunha** com topo preto chato

O vídeo foi produzido travado no molde **cunha**. Se a peça vendida for a cônica,
a produção precisa ser refeita com o outro conjunto de referências.

## Geração

Vídeo gerado com Seedance 2.5 em modo omni_reference, 1080p, 9:16, a partir das
imagens em `referencias/`. Montagem final por concatenação em ffmpeg:

```
ffmpeg -i bloco1.mp4 -i bloco2.mp4 -t 7.9 -i moto.mp4 -i bloco4.mp4 \
  -filter_complex "[0:v]setsar=1[v0];[1:v]setsar=1[v1];[2:v]setsar=1[v2];[3:v]setsar=1[v3];\
                   [v0][v1][v2][v3]concat=n=4:v=1:a=0[v]" \
  -map "[v]" -an -c:v libx264 -preset slow -crf 18 -pix_fmt yuv420p \
  -profile:v high -level 4.1 -r 24 -movflags +faststart saida.mp4
```
