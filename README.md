# Duelo Nacional de MCs — Estatísticas

Site estático (HTML + JavaScript, sem servidor) com estatísticas do Duelo Nacional de MCs:
estados & MCs, geral dos estados, ranking, chaveamentos por edição, confrontos entre estados,
campeões estaduais, recordes e uma aba Admin com inconsistências encontradas nos dados.

## Arquivos
- `index.html` — o site inteiro (layout + lógica)
- `data.js` — os dados já processados das planilhas
- `ferramentas/gerar_dados.py` — (opcional) regenera o `data.js` a partir das planilhas

## Atualizar os dados (opcional)
```
pip install openpyxl
python ferramentas/gerar_dados.py DUELO_NACIONALPOR_MC.csv Campeoes_Estaduais.xlsx data.js
```
Depois suba o novo `data.js` no GitHub (Add file → Upload files, substituindo o antigo).
