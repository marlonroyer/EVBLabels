# Gerador de Etiquetas - Herbário Evaldo Buttura (EVB)

Sistema web para geração de etiquetas botânicas a partir de planilhas Excel (.xls/.xlsx), com suporte a códigos de barras em imagens PNG.

## Funcionalidades

- Leitura de planilhas Excel (.xls e .xlsx)
- Upload de imagens de códigos de barras por lote
- Geração automática de etiquetas formatadas
- Exportação para PDF com layout A4 (2 etiquetas por linha)
- Interface moderna e responsiva

## Como usar

### Passo 1: Preparar os arquivos
- Tenha sua planilha Excel (.xls ou .xlsx) com as colunas necessárias (é a mesma planilha que foi usada para validar os dados no JABOT)
- Tenha as imagens dos códigos de barras em PNG (formato: `EVB00[numtombo].png`)
- As imagens são geradas a partir do [JABOTExtract](https://marlonroyer.github.io/JABOTExtract/)  que por sua vez precisa do PDF com os códigos de barra exportados diretamente pelo JABOT
- Ao final, as imagens dos códigos de barras devem seguir o padrão: EVB00[numtombo].png e todas elas devem estar juntas em uma mesma pasta

### Passo 2: Abrir o sistema
Acesse: [Gerador_Etiquetas_EVB](https://marlonroyer.github.io/Gerador_Etiquetas_EVB/)

### Passo 3: Carregar os dados
1. Selecione as imagens dos códigos de barras (.png)
2. Selecione a planilha de dados (.xls ou .xlsx)
3. As etiquetas serão geradas automaticamente

### Passo 4: Exportar
- Clique em "Gerar PDF" para baixar as etiquetas em formato A4
- Essa etapa pode ser rápida, como tambám pode ser muito lenta, depende do numero de etiquetas.

## Colunas necessárias na planilha

| Coluna na planilha | Descrição |
|-------------------|-----------|
| `numtombo` | Número de tombo (gera o EVB) |
| `family` | Família botânica |
| `genus` | Gênero |
| `sp1` | Epíteto específico |
| `author1` | Autor |
| `detby` | Determinador |
| `detdd` | Dia da determinação |
| `detmm` | Mês da determinação |
| `detyy` | Ano da determinação |
| `collector` | Coletor |
| `number` | Número de coleta |
| `projeto` | Projeto |

## Formato das imagens

As imagens dos códigos de barras devem seguir o padrão:
EVB00[numtombo].png
Exemplo: `EVB001234.png` (para o número de tombo 1234)

## Tecnologias utilizadas

- HTML5 / CSS3 / JavaScript
- [SheetJS](https://sheetjs.com/) - Leitura de arquivos Excel
- [jsPDF](https://github.com/parallax/jsPDF) - Geração de PDF
- [html2canvas](https://html2canvas.hertzen.com/) - Captura de HTML para imagem

## Licença

Este projeto é de uso interno do Herbário Evaldo Buttura.

## Desenvolvido por

Marlon Royer de Morais
