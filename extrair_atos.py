import os
import re
import time
from bs4 import BeautifulSoup
import pandas as pd
from playwright.sync_api import sync_playwright

# Instituições configuradas para extração
codigos = [96, 420]

todos_atos = []
arquivo_saida = "Atos_Instituicoes_MatoGrosso.xlsx"

print("=" * 80)
print("INICIANDO BOT DIÁRIO NA NUVEM")
print("=" * 80 + "\n")

with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=True,
        args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"]
    )
    page = browser.new_page()
    page.set_default_timeout(25000)

    for i, cod in enumerate(codigos, start=1):
        url = f"http://sigeduca.seduc.mt.gov.br/cee/wmceeinfinstensino.aspx?{cod}"
        print(f"[{i}/{len(codigos)}] Consultando instituição {cod}...")

        try:
            page.goto(url, timeout=35000)
            time.sleep(2)

            # Clica para abrir a aba 'ATOS DA INSTITUIÇÃO'
            botao = page.locator("input[name='BTATOSDAINSTITUICAO']")
            if botao.count() > 0:
                botao.click(force=True)
                time.sleep(2)
                if "ATO REGULATÓRIO" not in page.content().upper():
                    botao.click(force=True)
            time.sleep(4)

            # 1. Identifica Nome da Instituição
            soup_inicial = BeautifulSoup(page.content(), "html.parser")
            nome_escola = f"Instituição {cod}"
            linhas_geral = [l.strip() for l in soup_inicial.get_text(separator=" \n ").splitlines() if l.strip()]
            for idx, l in enumerate(linhas_geral):
                if "UNIDADE ESCOLAR" in l.upper() or "IDENTIFICAÇÃO DA MANTIDA" in l.upper():
                    cands = []
                    if ":" in l:
                        cands.append(l.split(":", 1)[-1].strip())
                    for off in [1, 2, 3]:
                        if idx + off < len(linhas_geral):
                            cands.append(linhas_geral[idx + off].strip())
                    for c in cands:
                        c_limpo = re.sub(r"^\d+\s*[-–]?\s*", "", c).strip()
                        c_limpo = re.split(r"MUNIC[ÍI]PIO", c_limpo, flags=re.IGNORECASE)[0].strip()
                        if c_limpo and not c_limpo.isdigit() and len(c_limpo) > 3:
                            nome_escola = c_limpo
                            break
                    if nome_escola != f"Instituição {cod}":
                        break

            print(f"   Nome identificado: {nome_escola}")

            # 2. Descobre total de páginas (se houver paginação)
            total_paginas = 1
            selects = page.locator("select")
            for s_idx in range(selects.count()):
                sel = selects.nth(s_idx)
                valores = sel.locator("option").all_inner_texts()
                num_opts = [v.strip() for v in valores if v.strip().isdigit()]
                if len(num_opts) > 1:
                    total_paginas = len(num_opts)
                    break

            print(f"   ✓ Total de páginas identificadas: {total_paginas}")

            # 3. Função de extração direta por bloco de texto
            def extrair_atos_pagina_atual(conteudo_html):
                soup = BeautifulSoup(conteudo_html, "html.parser")
                texto_pagina = soup.get_text("\n")
                atos_pag = []

                rotulos_principais = [
                    "Ato Regulatório", "Situação", "Município do Curso", "Curso",
                    "Modelo de Documento", "Tipo de Documento", "Nº Documento",
                    "Validade", "Data de Publicação", r"Pág\.", "Página da Publicação",
                    "Parecer", "Jornal de Circulação"
                ]

                blocos = re.split(r'(?i)(?=ATO\s+REGULAT[ÓO]RIO\s*:)', texto_pagina)

                for b in blocos:
                    if "ATO REGULATÓRIO:" not in b.upper():
                        continue

                    def pegar_campo(label, texto):
                        rotulos_sem_ele = [r for r in rotulos_principais if r.lower() != label.lower()]
                        padrao_fim = r"(?=(?:\b(?:" + "|".join(rotulos_sem_ele) + r")\s*:|\bAto\s*:|\bAto Antigo\s*:|$))"
                        if label.lower() == "curso":
                            p = rf"(?<!Munic[íi]pio do\s)(?<!Munic[íi]pio do )\bCurso\s*:\s*(.*?){padrao_fim}"
                        else:
                            p = rf"\b{label}\s*:\s*(.*?){padrao_fim}"

                        m = re.search(p, texto, flags=re.IGNORECASE | re.DOTALL)
                        if m:
                            lns = [l.strip() for l in m.group(1).strip().splitlines() if l.strip()]
                            return " ".join(lns)
                        return ""

                    curso_bruto = pegar_campo("Curso", b)
                    curso_limpo = re.sub(r"\s+", " ", curso_bruto).strip()

                    if not curso_limpo or curso_limpo in ["-", "--", "NÃO INFORMADO"]:
                        continue

                    # Filtra cursos técnicos / profissionalizantes
                    texto_curso = curso_limpo.upper()
                    if not ("TECNICO" in texto_curso or "TÉCNICO" in texto_curso or "PROFISSIONAL" in texto_curso):
                        continue

                    situacao = pegar_campo("Situação", b)
                    if "NÃO AUTORIZADO" in situacao.upper() or "NAO AUTORIZADO" in situacao.upper():
                        continue

                    ato_reg = pegar_campo("Ato Regulatório", b)

                    mun_curso = pegar_campo("Município do Curso", b)
                    mun_limpo = re.sub(r'[\r\n\t]+', ' ', mun_curso).strip().upper()
                    mun_limpo = re.sub(r'\s+', ' ', mun_limpo)
                    if not mun_limpo or mun_limpo in ["-", "--"]:
                        mun_limpo = "SINOP" if "SINOP" in b.upper() else "NÃO INFORMADO"

                    validade = pegar_campo("Validade", b)
                    data_pub = pegar_campo("Data de Publicação", b)

                    todas_datas = re.findall(r"\b\d{2}/\d{2}/\d{4}\b", validade)
                    if todas_datas:
                        data_fim_ref = todas_datas[-1]
                    else:
                        datas_pub = re.findall(r"\b\d{2}/\d{2}/\d{4}\b", data_pub)
                        data_fim_ref = datas_pub[-1] if datas_pub else "01/01/1900"

                    modelo_doc = pegar_campo("Modelo de Documento", b) or pegar_campo("Tipo de Documento", b)
                    num_doc = pegar_campo("Nº Documento", b)
                    pag = pegar_campo(r"Pág\.", b) or pegar_campo("Página da Publicação", b)

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
                        "_Data_Fim_Ref": data_fim_ref
                    }
                    atos_pag.append(item)

                return atos_pag

            # 4. Varredura por página
            for num_p in range(1, total_paginas + 1):
                if num_p > 1:
                    page.evaluate(f"""() => {{
                        const selects = Array.from(document.querySelectorAll('select'));
                        const sel = selects.find(s => Array.from(s.options).some(o => o.value == '{num_p}' || o.text == '{num_p}'));
                        if (sel) {{
                            sel.value = '{num_p}';
                            sel.dispatchEvent(new Event('change', {{ bubbles: true }}));
                        }}
                    }}""")
                    time.sleep(2.5)

                html_atual = page.content()
                atos_coletados = extrair_atos_pagina_atual(html_atual)
                print(f"      Página {num_p}: {len(atos_coletados)} atos capturados.")
                todos_atos.extend(atos_coletados)

            print(f"   >> Concluída coleta da instituição {cod}.\n")

        except Exception as e:
            print(f"   ✗ Erro no código {cod}: {e}\n")

    browser.close()

# 5. Processamento Incremental Consolidado
if todos_atos:
    df_novos = pd.DataFrame(todos_atos)
    df_novos["Data_Comparacao"] = pd.to_datetime(df_novos["_Data_Fim_Ref"], format="%d/%m/%Y", errors="coerce")

    if os.path.exists(arquivo_saida):
        try:
            df_existente = pd.read_excel(arquivo_saida)
            
            def extrair_data_val(val_txt):
                datas = re.findall(r"\b\d{2}/\d{2}/\d{4}\b", str(val_txt))
                return datas[-1] if datas else "01/01/1900"
                
            df_existente["_Data_Fim_Ref"] = df_existente["Validade"].apply(extrair_data_val)
            df_existente["Data_Comparacao"] = pd.to_datetime(df_existente["_Data_Fim_Ref"], format="%d/%m/%Y", errors="coerce")
            
            df_consolidado = pd.concat([df_existente, df_novos], ignore_index=True)
        except Exception:
            df_consolidado = df_novos
    else:
        df_consolidado = df_novos

    df_consolidado = df_consolidado.sort_values(
        by=["Código Instituição", "Município do Curso", "Curso", "Data_Comparacao"],
        ascending=[True, True, True, False]
    )

    df_final = df_consolidado.drop_duplicates(
        subset=["Código Instituição", "Município do Curso", "Curso"],
        keep="first"
    ).copy()

    colunas_limpar = ["_Data_Fim_Ref", "Data_Comparacao"]
    df_final = df_final.drop(columns=[c for c in colunas_limpar if c in df_final.columns])

    df_final.to_excel(arquivo_saida, index=False)
    print(f"✓ Base atualizada com {len(df_final)} cursos salvos em Excel.")
else:
    print("Nenhum ato novo encontrado.")
