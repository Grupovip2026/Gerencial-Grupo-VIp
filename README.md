# Dashboard Gerencial – Grupo VIP 2026

Dashboard financeiro hospedado em GitHub Pages, atualizado automaticamente via GitHub Actions.

## Estrutura do repositório

```
├── Resultado_Gerencial_2026.xlsx   ← planilha (você atualiza aqui)
├── scripts/
│   └── gerar_dados.py             ← lê o xlsx e gera docs/data.json
├── docs/
│   ├── index.html                 ← dashboard (GitHub Pages serve esta pasta)
│   └── data.json                  ← gerado automaticamente, não editar
└── .github/workflows/
    └── atualizar.yml              ← GitHub Actions
```

## Como atualizar os dados

### Opção 1 — Upload direto pelo GitHub (mais simples)
1. Abra o repositório em https://github.com/grupovip2026/Gerencial-Grupo-VIp
2. Clique no arquivo `Resultado_Gerencial_2026.xlsx`
3. Clique em **⋯ → Upload file** e selecione o arquivo novo
4. Clique em **Commit changes**
5. O GitHub Actions executa automaticamente e atualiza o dashboard em ~1 minuto

### Opção 2 — Git pelo terminal
```bash
git add Resultado_Gerencial_2026.xlsx
git commit -m "atualiza planilha maio/2026"
git push
```

### Atualização automática diária
O workflow roda todo dia às **06:00 BRT** (mesmo sem push), garantindo que o `data.json` reflita qualquer alteração recente.

### Execução manual
No GitHub: **Actions → Atualizar Dashboard → Run workflow**

## Configurar GitHub Pages (primeira vez)

1. Vá em **Settings → Pages**
2. Em **Source**, selecione `Deploy from a branch`
3. Branch: `main` | Folder: `/docs`
4. Salve — o site estará em `https://grupovip2026.github.io/Gerencial-Grupo-VIp/`

## Executar localmente (opcional)

```bash
pip install openpyxl
python scripts/gerar_dados.py
# abre docs/index.html no navegador
```
