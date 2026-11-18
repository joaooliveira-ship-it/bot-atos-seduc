import os
import re
import time
import requests
from bs4 import BeautifulSoup
import pandas as pd
from playwright.sync_api import sync_playwright
import sys  # <--- ADICIONE ESTA LINHA AQUI

# Instituições monitoradas (Pelegrino Cipriani + CIENTEC e demais ETECs)
TODOS_OS_CODIGOS = [96, 420]  # Adicione aqui os códigos das outras 15 escolas quando quiser

# Verifica se o GitHub Actions enviou um código específico ou pediu todas
parametro = sys.argv[1].strip() if len(sys.argv) > 1 else "TODAS"

if parametro.isdigit():
    codigos = [int(parametro)]  # Se enviou um código numérico (ex: 96), raspa apenas ele!
    print(f"🎯 Extração de Atos específica: código {codigos}")
else:
    codigos = TODOS_OS_CODIGOS  # Se for 'TODAS' ou agendamento diário, raspa todas
    print(f"🌐 Extração de Atos completa: varrendo todas as {len(codigos)} instituições.")

todos_atos = []
arquivo_saida = "Atos_Instituicoes_MatoGrosso.xlsx"

print("Iniciando extração com NAVEGAÇÃO REAL por todas as páginas (botão Próximo / AJAX)...\n")

with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=True,
        args=["--no-sandbox", "--disable-dev-shm-usage", "--disable-gpu"]
    )
    page = browser.new_page()

    for i, cod in enumerate(codigos, start=1):
        url = f"http://sigeduca.seduc.mt.gov.br/cee/wmceeinfinstensino.aspx?{cod}"
        print(f"[{i}/{len(codigos)}] Consultando instituição {cod}...")

        try:
            page.goto(url, timeout=35000)
            page.wait_for_load_state("domcontentloaded")
            time.sleep(2)

            botao = page.locator("input[name='BTATOSDAINSTITUICAO']")

            # Duplo clique para carregar a aba de atos
            botao.click()
            time.sleep(1.5)
            botao.click()
            page.wait_for_timeout(6000)

            # 1. Identifica Nome da Instituição
            soup_inicial = BeautifulSoup(page.content(), 'html.parser')
            nome_escola = "Não identificado"
            linhas_geral = [linha.strip() for linha in soup_inicial.get_text(separator=' \n ').splitlines() if linha.strip()]
            for idx, l in enumerate(linhas_geral):
                if "UNIDADE ESCOLAR" in l.upper() or "IDENTIFICAÇÃO DA MANTIDA" in l.upper():
                    cands = []
                    if ":" in l:
                        cands.append(l.split(":", 1)[-1].strip())
                    for off in [1, 2]:
                        if idx + off < len(linhas_geral):
                            cands.append(linhas_geral[idx + off].strip())
                    for c in cands:
                        c_limpo = re.sub(r'^\d+\s*[-–]?\s*', '', c).strip()
                        c_limpo = re.split(r'MUNIC[ÍI]PIO', c_limpo, flags=re.IGNORECASE)[0].strip()
                        if c_limpo and not c_limpo.isdigit() and len(c_limpo) > 3:
                            nome_escola = c_limpo
                            break
                    if nome_escola != "Não identificado":
                        break

            # 2. Descobre quantas páginas existem no select
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
                    print(f"   ✓ Instituição possui {total_paginas} páginas reais de atos.")
                    break

            # Função de extração de dados
            def extrair_atos_pagina_atual(conteudo_html):
                soup = BeautifulSoup(conteudo_html, 'html.parser')
                elementos = soup.find_all(['table', 'div'])
                atos_pag = []

                rotulos = [
                    "Ato Regulatório", "Situação", "Curso", "Município do Curso",
                    "Modelo de Documento", "Nº Documento", "Validade", "Data de Publicação",
                    r"Pág\.", "Parecer", "Ato"
                ]
                padrao_delimitador = r"(?=(?:" + "|".join(rotulos) + r")\s*:|$)"

                for elem in elementos:
                    txt = elem.get_text(separator='\n', strip=True)
                    if "ATO REGULATÓRIO:" in txt.upper() and "MODELO DE DOCUMENTO:" in txt.upper():
                        if "MOSTRAR ATOS:" in txt.upper() or txt.upper().count("ATO REGULATÓRIO:") > 1:
                            continue

                        def pegar_campo(nome_label, texto):
                            p = rf'{nome_label}\s*:\s*(.*?){padrao_delimitador}'
                            m = re.search(p, texto, flags=re.IGNORECASE | re.DOTALL)
                            if m:
                                lns = [l.strip() for l in m.group(1).strip().splitlines() if l.strip()]
                                return " ".join(lns)
                            return ""

                        curso_bruto = pegar_campo("Curso", txt)
                        curso_limpo = re.sub(r'\s+', ' ', curso_bruto).strip()

                        # Descarta se não houver nome de curso
                        if not curso_limpo or curso_limpo in ["-", "--", "NÃO INFORMADO"]:
                            continue

                        # Mantém Cursos Técnicos / Educação Profissional
                        texto_curso = curso_limpo.upper()
                        if not ("TECNICO" in texto_curso or "TÉCNICO" in texto_curso or "PROFISSIONAL" in texto_curso):
                            continue

                        situacao = pegar_campo("Situação", txt)
                        if "NÃO AUTORIZADO" in situacao.upper() or "NAO AUTORIZADO" in situacao.upper():
                            continue

                        ato_reg = pegar_campo("Ato Regulatório", txt)

                        # Município do Curso
                        mun_curso = pegar_campo("Município do Curso", txt)
                        mun_limpo = re.sub(r'[\r\n\t]+', ' ', mun_curso).strip().upper()
                        mun_limpo = re.sub(r'\s+', ' ', mun_limpo)
                        if not mun_limpo or mun_limpo in ["-", "--"]:
                            mun_limpo = "CUIABÁ" if cod == 96 else "SINOP"

                        validade = pegar_campo("Validade", txt)
                        data_pub = pegar_campo("Data de Publicação", txt)

                        todas_datas = re.findall(r'\b\d{2}/\d{2}/\d{4}\b', validade)
                        if todas_datas:
                            data_fim_ref = todas_datas[-1]
                        else:
                            datas_pub = re.findall(r'\b\d{2}/\d{2}/\d{4}\b', data_pub)
                            data_fim_ref = datas_pub[-1] if datas_pub else "01/01/1900"

                        modelo_doc = pegar_campo("Modelo de Documento", txt)
                        num_doc = pegar_campo("Nº Documento", txt)
                        pag = pegar_campo(r"Pág\.", txt)

                        item = {
                            'Código Instituição': cod,
                            'Nome da Instituição': nome_escola,
                            'Município do Curso': mun_limpo,
                            'Curso': curso_limpo,
                            'Ato Regulatório': ato_reg,
                            'Situação': situacao,
                            'Modelo de Documento': modelo_doc,
                            'Nº Documento': num_doc,
                            'Validade': validade,
                            'Data de Publicação': data_pub,
                            'Pág.': pag,
                            '_Data_Fim_Ref': data_fim_ref
                        }
                        atos_pag.append(item)

                return atos_pag

            # 3. Varredura com avanço real página por página
            for num_p in range(1, total_paginas + 1):
                print(f"   -> Extraindo Página {num_p} de {total_paginas}...")

                if num_p > 1:
                    btn_proximo = page.locator("a:has-text('Próximo'), a:has-text('Proximo'), span:has-text('Próximo')")
                    if btn_proximo.count() > 0 and btn_proximo.first.is_visible():
                        btn_proximo.first.click()
                    elif select_pag:
                        select_pag.select_option(str(num_p))
                        select_pag.dispatch_event("change")

                    page.wait_for_timeout(4000)

                html_atual = page.content()
                atos_coletados = extrair_atos_pagina_atual(html_atual)

                if atos_coletados:
                    doc_atual = atos_coletados[0]['Nº Documento']
                    mun_atual = atos_coletados[0]['Município do Curso']
                    print(f"      ✓ {len(atos_coletados)} atos capturados (Ex: {mun_atual} - Doc: {doc_atual})")
                else:
                    print("      - Nenhum curso técnico nesta página.")

                todos_atos.extend(atos_coletados)

            print(f"\n   >> Coleta concluída para escola {cod}: {len(todos_atos)} registros brutos.\n")

        except Exception as e:
            print(f"   ✗ Erro no código {cod}: {e}\n")

    browser.close()

# 4. FILTRAGEM: ATO MAIS RECENTE POR (MUNICÍPIO + CURSO TÉCNICO)
if todos_atos:
    df = pd.DataFrame(todos_atos)

    # Remove duplicatas brutas idênticas
    df = df.drop_duplicates(subset=['Código Instituição', 'Município do Curso', 'Curso', 'Nº Documento', 'Validade'])
    print(f"Total bruto de atos técnicos capturados: {len(df)}")

    # Converte data para ordenação cronológica
    df['Data_Comparacao'] = pd.to_datetime(df['_Data_Fim_Ref'], format='%d/%m/%Y', errors='coerce')

    # Ordena da maior validade (mais recente) para a menor
    df = df.sort_values(
        by=['Código Instituição', 'Município do Curso', 'Curso', 'Data_Comparacao'],
        ascending=[True, True, True, False]
    )

    # Mantém SOMENTE o ato mais recente de cada par (Município + Curso Técnico)
    df_mais_recentes = df.drop_duplicates(
        subset=['Código Instituição', 'Município do Curso', 'Curso'],
        keep='first'
    ).copy()

    df_final = df_mais_recentes.drop(columns=['_Data_Fim_Ref', 'Data_Comparacao'])

    # Ordena as colunas
    colunas_ordenadas = [
        'Código Instituição', 'Nome da Instituição', 'Município do Curso', 'Curso',
        'Ato Regulatório', 'Situação', 'Modelo de Documento', 'Nº Documento',
        'Validade', 'Data de Publicação', 'Pág.'
    ]
    df_final = df_final[[c for c in colunas_ordenadas if c in df_final.columns]]

    # 1. Salva no arquivo Excel padrão (backup no GitHub)
    df_final.to_excel(arquivo_saida, index=False)
    print("=" * 80)
    print(f"✓ Sucesso! {len(df_final)} cursos técnicos únicos gravados em '{arquivo_saida}'.")
    print("=" * 80)

    # 2. Envia para o Google Planilhas na Web (se a URL estiver configurada)
    url_webhook = os.getenv("URL_GOOGLE_PLANILHA")
    if url_webhook:
        try:
            print("Enviando dados para a planilha do Google na Web...")
            linhas = [df_final.columns.tolist()] + df_final.fillna("").astype(str).values.tolist()
            resp = requests.post(url_webhook, json=linhas, timeout=30)
            if resp.status_code == 200:
                print("✓ Sucesso: Google Planilhas atualizado na Web com sucesso!")
            else:
                print(f"⚠️ Aviso ao enviar para o Google: {resp.text}")
        except Exception as err:
            print(f"⚠️ Não foi possível sincronizar com o Google: {err}")
    else:
        print("ℹ️ URL_GOOGLE_PLANILHA não configurada nos Secrets. Dados salvos apenas no Excel.")

    # Resumo de todas as cidades encontradas
    cidades = sorted(df_final['Município do Curso'].unique())
    print(f"\nTotal de cidades/polos identificados: {len(cidades)}")
    print("Cidades capturadas:")
    for cid in cidades:
        qtd = len(df_final[df_final['Município do Curso'] == cid])
        print(f" - {cid} ({qtd} cursos)")

else:
    print("Nenhum curso técnico encontrado.")
