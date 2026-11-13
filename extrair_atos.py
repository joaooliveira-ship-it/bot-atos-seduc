import os
import re
import time
from bs4 import BeautifulSoup
import pandas as pd
from playwright.sync_api import sync_playwright

# Instituições configuradas
codigos = [96, 420]

todos_atos = []
arquivo_saida = "Atos_Instituicoes_MatoGrosso.xlsx"

print("=" * 80)
print("INICIANDO EXTRAÇÃO CONSOLIDADA (CIPRIANI + CIENTEC)")
print("=" * 80 + "\n")

with sync_playwright() as p:
  browser = p.chromium.launch(
      headless=True,
      args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"],
  )
  page = browser.new_page()
  page.set_default_timeout(45000)

  for i, cod in enumerate(codigos, start=1):
    url = f"http://sigeduca.seduc.mt.gov.br/cee/wmceeinfinstensino.aspx?{cod}"
    print(f"[{i}/{len(codigos)}] Consultando instituição {cod}...")

    try:
      page.goto(url, timeout=50000)
      time.sleep(3)

      # 1. Identifica Nome da Instituição na página inicial
      soup_inicial = BeautifulSoup(page.content(), "html.parser")
      nome_escola = f"Instituição {cod}"
      linhas_geral = [
          l.strip()
          for l in soup_inicial.get_text(separator=" \n ").splitlines()
          if l.strip()
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
            )[0].strip()
            if c_limpo and not c_limpo.isdigit() and len(c_limpo) > 3:
              nome_escola = c_limpo
              break
          if nome_escola != f"Instituição {cod}":
            break

      print(f"   Nome identificado: {nome_escola}")

      # 2. Clica para abrir a aba 'ATOS DA INSTITUIÇÃO' e aguarda com segurança
      botao = page.locator("input[name='BTATOSDAINSTITUICAO']")
      if botao.count() > 0:
        botao.click(force=True)
        time.sleep(3)

      # Garante que a tabela de atos realmente carregou
      for tentativa in range(5):
        if "ATO REGULATÓRIO" in page.content().upper():
          break
        if botao.count() > 0:
          botao.click(force=True)
        time.sleep(3)

      if "ATO REGULATÓRIO" not in page.content().upper():
        print(f"   ✗ Não foi possível carregar a grade de atos da escola {cod}.")
        continue

      # 3. Identifica total de páginas
      total_paginas = 1
      select_pag = page.locator("select").first
      if select_pag.count() > 0:
        opts = select_pag.locator("option").all_inner_texts()
        nums = [o.strip() for o in opts if o.strip().isdigit()]
        if len(nums) > 1:
          total_paginas = len(nums)

      print(f"   ✓ Total de páginas identificadas: {total_paginas}")

      # 4. Função de extração de cada ato
      def extrair_atos(conteudo_html):
        soup = BeautifulSoup(conteudo_html, "html.parser")
        texto_pagina = soup.get_text("\n")
        atos_pag = []

        rotulos = [
            "Ato Regulatório",
            "Situação",
            "Município do Curso",
            "Curso",
            "Modelo de Documento",
            "Tipo de Documento",
            "Nº Documento",
            "Validade",
            "Data de Publicação",
            r"Pág\.",
            "Página da Publicação",
            "Parecer",
            "Jornal de Circulação",
        ]

        blocos = re.split(r"(?i)(?=ATO\s+REGULAT[ÓO]RIO\s*:)", texto_pagina)

        for b in blocos:
          if "ATO REGULATÓRIO:" not in b.upper():
            continue

          def pegar_campo(label, texto):
            rotulos_sem_ele = [r for r in rotulos if r.lower() != label.lower()]
            padrao_fim = (
                r"(?=(?:\b(?:"
                + "|".join(rotulos_sem_ele)
                + r")\s*:|\bAto\s*:|\bAto Antigo\s*:|$))"
            )
            if label.lower() == "curso":
              p = (
                  rf"(?<!Munic[íi]pio do\s)(?<!Munic[íi]pio do"
                  rf" )\bCurso\s*:\s*(.*?){padrao_fim}"
              )
            else:
              p = rf"\b{label}\s*:\s*(.*?){padrao_fim}"

            m = re.search(p, texto, flags=re.IGNORECASE | re.DOTALL)
            if m:
              lns = [
                  l.strip()
                  for l in m.group(1).strip().splitlines()
                  if l.strip()
              ]
              return " ".join(lns)
            return ""

          curso_bruto = pegar_campo("Curso", b)
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

          situacao = pegar_campo("Situação", b)
          if (
              "NÃO AUTORIZADO" in situacao.upper()
              or "NAO AUTORIZADO" in situacao.upper()
          ):
            continue

          ato_reg = pegar_campo("Ato Regulatório", b)

          mun_curso = pegar_campo("Município do Curso", b)
          mun_limpo = re.sub(r"[\r\n\t]+", " ", mun_curso).strip().upper()
          mun_limpo = re.sub(r"\s+", " ", mun_limpo)
          if not mun_limpo or mun_limpo in ["-", "--"]:
            mun_limpo = (
                "CUIABÁ"
                if cod == 96
                else ("SINOP" if cod == 420 else "NÃO INFORMADO")
            )

          validade = pegar_campo("Validade", b)
          data_pub = pegar_campo("Data de Publicação", b)

          todas_datas = re.findall(r"\b\d{2}/\d{2}/\d{4}\b", validade)
          if todas_datas:
            data_fim_ref = todas_datas[-1]
          else:
            datas_pub = re.findall(r"\b\d{2}/\d{2}/\d{4}\b", data_pub)
            data_fim_ref = datas_pub[-1] if datas_pub else "01/01/1900"

          modelo_doc = pegar_campo("Modelo de Documento", b) or pegar_campo(
              "Tipo de Documento", b
          )
          num_doc = pegar_campo("Nº Documento", b)
          pag = pegar_campo(r"Pág\.", b) or pegar_campo(
              "Página da Publicação", b
          )

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

      # 5. Varredura das páginas
      atos_escola_atual = []
      for num_p in range(1, total_paginas + 1):
        if num_p > 1:
          seletor = page.locator("select").first
          if seletor.count() > 0:
            seletor.select_option(str(num_p))
            page.evaluate("""() => {
                            const sel = document.querySelector('select');
                            if (sel) {
                                sel.dispatchEvent(new Event('change', { bubbles: true }));
                                if (typeof gx !== 'undefined' && gx.evt) {
                                    gx.evt.onchange(sel);
                                }
                            }
                        }""")
            time.sleep(3.5)

        html_atual = page.content()
        atos_pag = extrair_atos(html_atual)
        atos_escola_atual.extend(atos_pag)

      print(
          f"   >> Concluída instituição {cod}: {len(atos_escola_atual)} atos"
          " técnicos encontrados no total.\n"
      )
      todos_atos.extend(atos_escola_atual)

    except Exception as e:
      print(f"   ✗ Erro ao processar instituição {cod}: {e}\n")

  browser.close()

# 6. Gravação Final Consolidada (Cipriani + CIENTEC)
if todos_atos:
  df_novos = pd.DataFrame(todos_atos)
  df_novos["Data_Comparacao"] = pd.to_datetime(
      df_novos["_Data_Fim_Ref"], format="%d/%m/%Y", errors="coerce"
  )

  # Ordena pela data mais recente
  df_consolidado = df_novos.sort_values(
      by=[
          "Código Instituição",
          "Município do Curso",
          "Curso",
          "Data_Comparacao",
      ],
      ascending=[True, True, True, False],
  )

  # Mantém o ato mais recente de cada curso por instituição
  df_final = df_consolidado.drop_duplicates(
      subset=["Código Instituição", "Município do Curso", "Curso"], keep="first"
  ).copy()

  colunas_limpar = ["_Data_Fim_Ref", "Data_Comparacao"]
  df_final = df_final.drop(
      columns=[c for c in colunas_limpar if c in df_final.columns]
  )

  df_final.to_excel(arquivo_saida, index=False)
  print("=" * 80)
  print(
      f"✓ SUCESSO: Planilha gerada com {len(df_final)} cursos salvos em"
      f" '{arquivo_saida}'!"
  )
  print(f"Instituições presentes na planilha: {df_final['Código Instituição'].unique()}")
  print("=" * 80)
else:
  print("Nenhum ato novo encontrado.")
