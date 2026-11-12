import os
import re
import time
from bs4 import BeautifulSoup
import pandas as pd
from playwright.sync_api import sync_playwright

# Lista de instituições que você deseja monitorar
codigos = [96]

todos_atos = []
arquivo_saida = "Atos_Instituicoes_MatoGrosso.xlsx"

print("=" * 80)
print("INICIANDO BOT DIÁRIO NA NUVEM (GITHUB ACTIONS)")
print("=" * 80 + "\n")

with sync_playwright() as p:
  browser = p.chromium.launch(headless=True)
  page = browser.new_page()

  for i, cod in enumerate(codigos, start=1):
    url = f"http://sigeduca.seduc.mt.gov.br/cee/wmceeinfinstensino.aspx?{cod}"
    print(f"[{i}/{len(codigos)}] Consultando instituição {cod}...")

    try:
      page.goto(url, timeout=45000)
      page.wait_for_load_state("domcontentloaded")
      time.sleep(2)

      botao = page.locator("input[name='BTATOSDAINSTITUICAO']")
      botao.click()
      time.sleep(1.5)
      botao.click()
      page.wait_for_timeout(6000)

      # 1. Identifica Nome da Instituição
      soup_inicial = BeautifulSoup(page.content(), "html.parser")
      nome_escola = "Não identificado"
      linhas_geral = [
          linha.strip()
          for linha in soup_inicial.get_text(separator=" \n ").splitlines()
          if linha.strip()
      ]
      for idx, l in enumerate(linhas_geral):
        if (
            "UNIDADE ESCOLAR" in l.upper()
            or "IDENTIFICAÇÃO DA MANTIDA" in l.upper()
        ):
          cands = []
          if ":" in l:
            cands.append(l.split(":", 1)[-1].strip())
          for off in [1, 2, 3]:
            if idx + off < len(linhas_geral):
              cands.append(linhas_geral[idx + off].strip())
          for c in cands:
            c_limpo = re.sub(r"^\d+\s*[-–]?\s*", "", c).strip()
            c_limpo = re.split(
                r"MUNIC[ÍI]PIO", c_limpo, flags=re.IGNORECASE
            ).strip()
            if c_limpo and not c_limpo.isdigit() and len(c_limpo) > 3:
              nome_escola = c_limpo
              break
          if nome_escola != "Não identificado":
            break

      # 2. Total de páginas
      total_paginas = 1
      select_pag = None
      selects = page.locator("select")
      for s_idx in range(selects.count()):
        sel = selects.nth(s_idx)
        valores = sel.locator("option").all_inner_texts()
        num_opts = [v.strip() for v in valores if v.strip().isdigit()]
        if len(num_opts) > 1:
          select_pag = sel
          total_paginas = len(num_opts)
          print(f"   ✓ Instituição possui {total_paginas} páginas de atos.")
          break

      # Função de extração
      def extrair_atos_pagina_atual(conteudo_html):
        soup = BeautifulSoup(conteudo_html, "html.parser")
        elementos = soup.find_all(["table", "div"])
        atos_pag = []

        rotulos = [
            "Ato Regulatório",
            "Situação",
            "Curso",
            "Município do Curso",
            "Modelo de Documento",
            "Nº Documento",
            "Validade",
            "Data de Publicação",
            r"Pág\.",
            "Parecer",
            "Ato",
        ]
        padrao_delimitador = r"(?=(?:" + "|".join(rotulos) + r")\s*:|$)"

        for elem in elementos:
          txt = elem.get_text(separator="\n", strip=True)
          if (
              "ATO REGULATÓRIO:" in txt.upper()
              and "MODELO DE DOCUMENTO:" in txt.upper()
          ):
            if (
                "MOSTRAR ATOS:" in txt.upper()
                or txt.upper().count("ATO REGULATÓRIO:") > 1
            ):
              continue

            def pegar_campo(nome_label, texto):
              p = rf"{nome_label}\s*:\s*(.*?){padrao_delimitador}"
              m = re.search(p, texto, flags=re.IGNORECASE | re.DOTALL)
              if m:
                lns = [
                    l.strip()
                    for l in m.group(1).strip().splitlines()
                    if l.strip()
                ]
                return " ".join(lns)
              return ""

            curso_bruto = pegar_campo("Curso", txt)
            curso_limpo = re.sub(r"\s+", " ", curso_bruto).strip()

            if not curso_limpo or curso_limpo in ["-", "--", "NÃO INFORMADO"]:
              continue

            texto_curso = curso_limpo.upper()
            if not (
                "TECNICO" in texto_curso
                or "TÉCNICO" in texto_curso
                or "PROFISSIONAL" in texto_curso
            ):
              continue

            situacao = pegar_campo("Situação", txt)
            if (
                "NÃO AUTORIZADO" in situacao.upper()
                or "NAO AUTORIZADO" in situacao.upper()
            ):
              continue

            ato_reg = pegar_campo("Ato Regulatório", txt)

            mun_curso = pegar_campo("Município do Curso", txt)
            mun_limpo = re.sub(r"[\r\n\t]+", " ", mun_curso).strip().upper()
            mun_limpo = re.sub(r"\s+", " ", mun_limpo)
            if not mun_limpo or mun_limpo in ["-", "--"]:
              mun_limpo = "CUIABÁ"

            validade = pegar_campo("Validade", txt)
            data_pub = pegar_campo("Data de Publicação", txt)

            todas_datas = re.findall(r"\b\d{2}/\d{2}/\d{4}\b", validade)
            if todas_datas:
              data_fim_ref = todas_datas[-1]
            else:
              datas_pub = re.findall(r"\b\d{2}/\d{2}/\d{4}\b", data_pub)
              data_fim_ref = datas_pub[-1] if datas_pub else "01/01/1900"

            modelo_doc = pegar_campo("Modelo de Documento", txt)
            num_doc = pegar_campo("Nº Documento", txt)
            pag = pegar_campo(r"Pág\.", txt)

            item = {
                "Código Instituição": cod,
                "Nome da Instituição": nome_escola,
                "Município do Curso": mun_limpo,
                "Curso": curso_limpo,
                "Ato Regulatório": ato_reg,
                "Situação": situacao,
                "Modelo de Documento": modelo_doc,
                "Nº Documento": num_doc,
                "Validade": validade,
                "Data de Publicação": data_pub,
                "Pág.": pag,
                "_Data_Fim_Ref": data_fim_ref,
            }
            atos_pag.append(item)

        return atos_pag

      # 3. Navegação pelas páginas
      for num_p in range(1, total_paginas + 1):
        if num_p > 1:
          btn_proximo = page.locator(
              "a:has-text('Próximo'), a:has-text('Proximo'),"
              " span:has-text('Próximo')"
          )
          if btn_proximo.count() > 0 and btn_proximo.first.is_visible():
            btn_proximo.first.click()
          elif select_pag:
            select_pag.select_option(str(num_p))
            select_pag.dispatch_event("change")
          page.wait_for_timeout(3500)

        html_atual = page.content()
        atos_coletados = extrair_atos_pagina_atual(html_atual)
        todos_atos.extend(atos_coletados)

      print(f"   >> Concluída coleta da instituição {cod}.")

    except Exception as e:
      print(f"   ✗ Erro no código {cod}: {e}")

  browser.close()

# 4. Processamento Incremental (Mantém o ato mais recente de cada curso por município)
if todos_atos:
  df_novos = pd.DataFrame(todos_atos)
  df_novos["Data_Comparacao"] = pd.to_datetime(
      df_novos["_Data_Fim_Ref"], format="%d/%m/%Y", errors="coerce"
  )

  if os.path.exists(arquivo_saida):
    try:
      df_existente = pd.read_excel(arquivo_saida)

      def extrair_data_val(val_txt):
        datas = re.findall(r"\b\d{2}/\d{2}/\d{4}\b", str(val_txt))
        return datas[-1] if datas else "01/01/1900"

      df_existente["_Data_Fim_Ref"] = df_existente["Validade"].apply(
          extrair_data_val
      )
      df_existente["Data_Comparacao"] = pd.to_datetime(
          df_existente["_Data_Fim_Ref"], format="%d/%m/%Y", errors="coerce"
      )

      df_consolidado = pd.concat([df_existente, df_novos], ignore_index=True)
    except Exception:
      df_consolidado = df_novos
  else:
    df_consolidado = df_novos

  df_consolidado = df_consolidado.sort_values(
      by=[
          "Código Instituição",
          "Município do Curso",
          "Curso",
          "Data_Comparacao",
      ],
      ascending=[True, True, True, False],
  )

  df_final = df_consolidado.drop_duplicates(
      subset=["Código Instituição", "Município do Curso", "Curso"], keep="first"
  ).copy()

  colunas_limpar = ["_Data_Fim_Ref", "Data_Comparacao"]
  df_final = df_final.drop(
      columns=[c for c in colunas_limpar if c in df_final.columns]
  )

  df_final.to_excel(arquivo_saida, index=False)
  print(f"✓ Base atualizada com {len(df_final)} cursos salvos em Excel.")
else:
  print("Nenhum ato novo encontrado.")
No canto superior direito, clique no botão verde Commit changes... e confirme no popup.
Passo 2: Criar o agendamento (.github/workflows/robo_diario.yml)
Clique novamente em Add file > Create new file.
No campo name your file..., cole exatamente este texto (com as barras):
.github/workflows/robo_diario.yml
(Ao digitar a barra /, o GitHub criará as pastas automaticamente!)
Na caixa de texto grande abaixo, cole a configuração do agendamento diário:
name: Rotina Diaria - Extracao de Atos SEDUC-MT

on:
  schedule:
    # Roda todos os dias às 10:00 UTC (06:00 da manhã no horário de Mato Grosso)
    - cron: '0 10 * * *'
  workflow_dispatch: # Permite rodar manualmente clicando em um botão

jobs:
  executar-bot:
    runs-on: ubuntu-latest

    permissions:
      contents: write # Permissão para salvar o Excel de volta no repositório

    steps:
      - name: Baixar arquivos do repositório
        uses: actions/checkout@v4

      - name: Configurar o Python
        uses: actions/setup-python@v5
        with:
          python-version: '3.11'

      - name: Instalar bibliotecas
        run: |
          python -m pip install --upgrade pip
          pip install -r requirements.txt
          playwright install chromium --with-deps

      - name: Executar o Robô de Extração
        run: |
          python extrair_atos.py

      - name: Salvar planilha atualizada no GitHub
        run: |
          git config --global user.name "Robo SEDUC Bot"
          git config --global user.email "bot@githubactions.local"
          git add Atos_Instituicoes_MatoGrosso.xlsx
          git diff --quiet && git diff --staged --quiet || (git commit -m "Auto: Planilha atualizada [skip ci]" && git push)
