# 🐔 Painel da Granja

Gestão simples e bonita para quem cria **galinhas poedeiras**. Um app de Windows
que registra ovos, ração, gastos, vendas e mostra o **lucro ou prejuízo do mês** —
e pode abrir sozinho toda vez que você ligar o computador.

## ⬇️ Baixar (não precisa instalar Python)

Vá em **[Releases](../../releases)** e baixe o arquivo
`Instalador - Painel da Granja.exe`. Dê 2 cliques e siga as telas em português.

## ✨ O que ele faz

| Aba | Função |
|---|---|
| 📝 **Registro do dia** | Ovos coletados, dúzias vendidas × preço da dúzia, quebrados, mortes, ração usada e gastos — com receita calculada ao vivo |
| 🌾 **Ração** | Estoque atual, consumo médio/dia e **alerta vermelho quando falta menos de 7 dias** |
| 💰 **Lucro / Prejuízo** | Resumo do mês, custo por ovo, média por galinha/dia e **exportação em Excel (.xlsx) formatado** |
| ☁️ **Clima** | Previsão pela sua cidade (Open-Meteo), com alertas de frio forte e calor extremo |
| ⚙️ **Ajustes** | Quantidade de galinhas, estoque inicial de ração, acesso à pasta de dados |

## 💻 Instalar do código (para desenvolvedores)

```bash
pip install openpyxl
python app.py
```

Para gerar o EXE:

```bash
pip install pyinstaller openpyxl
pyinstaller --onefile --windowed --name PainelGranja --icon icone.ico app.py
```

## 🗂️ Dados

Tudo é gravado em `%APPDATA%\PainelGranja` (SQLite). Ao desinstalar,
**os dados não são apagados**. Copie essa pasta de vez em quando para um pen drive.

## 📄 Licença

[MIT](LICENSE)

## 🖼️ Créditos da imagem

Ícone: foto ["Chicken"](https://www.flickr.com/photos/29278394@N00/8109611989)
de **normanack**, licença [CC BY 2.0](https://creativecommons.org/licenses/by/2.0/)
(via [Openverse](https://openverse.org)).
