
# Excel → App Web (Streamlit)

App web para **editar**, **gerar relatórios** e **criar gráficos** a partir de **todas as abas** de um arquivo Excel (`.xlsm/.xlsx`).  
Salva de volta em **.xlsm preservando macros**.

## Funcionalidades
- 📝 **Edição:** tabela editável por aba (adicionar/excluir linhas).
- 📊 **Relatórios:** agrupamentos + agregações (sum/mean/median/min/max/count) e download CSV.
- 📈 **Gráficos:** linha, barra, área e dispersão (Altair).
- 💾 **Salvar/Exportar:** salva `.xlsm` com macros (`keep_vba=True`) e exporta todas as abas em ZIP de CSVs.

---

## Rodar localmente
```bash
pip install -r requirements.txt
streamlit run app.py
```
> Coloque seu `Strategic_Plan_2025_Rev01.xlsm` na **mesma pasta** do `app.py` ou faça **upload** pela barra lateral ao abrir o app.

---

## Deploy (3 jeitos)

### A) Streamlit Community Cloud (rápido)
1. Crie um repositório no GitHub contendo: `app.py`, `requirements.txt`, `README.md` (e **opcional** o `.xlsm`).
2. Vá em https://streamlit.io → **Deploy an app** → conecte ao seu repositório.
3. Selecione `app.py` como entrypoint e faça o deploy.

> Se não quiser subir o `.xlsm`, use o **upload** no app quando abrir.

### B) Hugging Face Spaces
1. Crie um **Space** do tipo **Streamlit**.
2. Envie `app.py`, `requirements.txt` e `README.md` (e opcionalmente o `.xlsm`).
3. O build roda automático.

### C) Render (com persistência de arquivo)
1. Crie um **Web Service** e adicione **Disco Persistente**.
2. Use o `Procfile` deste repo (porta `$PORT` é injetada pelo Render).
3. Faça upload do `.xlsm` para a pasta persistente do serviço.

---

## Docker (opcional)
```bash
docker build -t excel-webapp .
docker run -p 8501:8501 -v %cd%:/app excel-webapp   # Windows (PowerShell)
# ou
docker run -p 8501:8501 -v $(pwd):/app excel-webapp  # Linux/macOS
```
> Com o volume `-v`, o app enxerga os arquivos locais (incluindo seu `.xlsm`).

---

## Estrutura
```
.
├── app.py
├── requirements.txt
├── README.md
├── .gitignore
├── Procfile
└── Dockerfile
```

---

## Notas
- Formatação/estilos do Excel podem ser perdidos na reescrita dos dados (os **valores** são mantidos).
- Colunas de data e numéricas têm conversão automática básica após edição.
- Se precisar **relatórios/gráficos fixos** (pré-configurados) ou **regras de edição** por coluna, abra uma _issue_ ou me diga os detalhes que eu ajusto.
