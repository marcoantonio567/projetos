# projetos_rpa / projetos

Collection of small RPA-style automation exercises in Python (spreadsheets, documents, images, and browser automation).

## Structure

- `projeto1/`: reads `fornecedores.xlsx` and generates `.docx` contracts (Word)
- `projeto2/`: reads `planilha_alunos.xlsx` and generates certificate images from a template
- `projeto3/`: reads `produtos.xlsx` and fills a local HTML form via Selenium
- `projeto4/`: scrapes products/prices from a website and saves them to `produtos.xlsx`
- `projeto5/`: generates fake data into Excel and renders text onto an image

## Prerequisites

- Python 3.10+ recommended
- For Selenium projects (`projeto3`, `projeto4`):
  - Firefox installed (scripts use `webdriver.Firefox`)
  - `selenium` installed
  - Driver availability depends on your Selenium setup (Selenium Manager may download drivers automatically; otherwise ensure `geckodriver` is available in PATH)

## Setup

From this `projetos/` folder:

```bash
python -m venv .venv
.\.venv\Scripts\activate
pip install -r requirements.txt
```

Some projects require extra dependencies not listed in `requirements.txt`:

```bash
pip install pillow selenium faker phonenumbers
```

## Running

Examples (from `projetos/`):

```bash
python projeto1\app.py
python projeto2\app.py
python projeto3\app.py
python projeto4\app.py
python projeto5\principal.py
python projeto5\gerar_dados_ficticios.py
```

## Notes / Common fixes

- Paths: some scripts use hard-coded absolute paths. If you move the repository, update those paths to relative ones.
- `projeto1`: creates files under `projeto1\contratos\`. Create that folder before running.
- `projeto2`: generates files into `projeto2\certificados\` (created automatically).
- `projeto3`: expects a local URL like `http://127.0.0.1:3000/projeto3/sistema.html`.
  - One simple way to serve the folder:

```bash
python -m http.server 3000
```

- Fonts/images: `projeto2` and `projeto5` rely on `.ttf` files and template images. Keep them in the same relative paths used by the scripts.
