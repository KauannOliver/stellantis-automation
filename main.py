import os
import time
from datetime import datetime, timedelta, date

import pandas as pd
import win32com.client as win32

def refresh_excel_workbook(excel_path):
    ### ab: abri excel, atual cnxs e salva
    print("iniciando excel para atualizar a planilha...")
    excel_app = win32.Dispatch("Excel.Application")
    excel_app.Visible = True
    try:
        wb = excel_app.Workbooks.Open(excel_path)
        ### ab: atual consultas
        wb.RefreshAll()
        ### ab: espera 180s
        time.sleep(180)
        ### ab: salva e fecha
        wb.Save()
        wb.Close()
    except Exception as e:
        print("erro ao atualizar excel:", e)
    finally:
        excel_app.Quit()
    print("planilha atualizada e salva com sucesso!")

def next_monday(dt: date) -> date:
    ### ab: ret prox segunda
    while dt.weekday() != 0:
        dt += timedelta(days=1)
    return dt

def get_final_due_date(data_emissao: date, lead_time: int, dtl_descri: str) -> date:
    ### ab: calc data final somando lead_time e mais 30 ou 50 se spot
    base_date = data_emissao + timedelta(days=lead_time)
    dtl_lower = str(dtl_descri).lower()
    if "spot" in dtl_lower:
        base_date += timedelta(days=50)
    else:
        base_date += timedelta(days=30)
    ### ab: prox segunda
    final_date = next_monday(base_date)
    return final_date

def main():
    ### ab: caminhos e arquivo
    pasta_base = r"S:\Publico\Dep. Financeiro - Faturamento\2. Stellantis"
    nome_planilha_base = "00. Base Stellantis.xlsx"
    caminho_arquivo = os.path.join(pasta_base, nome_planilha_base)

    ### ab: (opc) atual plan
    # refresh_excel_workbook(caminho_arquivo)

    ### ab: le planilhas
    df_stell = pd.read_excel(caminho_arquivo, sheet_name="queryStellantis")
    df_faturados = pd.read_excel(caminho_arquivo, sheet_name="queryFaturados")

    ### ab: ajusta data emissao
    if not pd.api.types.is_datetime64_any_dtype(df_stell["DATA EMISSAO"]):
        df_stell["DATA EMISSAO"] = pd.to_datetime(df_stell["DATA EMISSAO"])

    ### ab: listas de saida
    a_faturar_data = []
    verificar_data = []

    ### ab: agrupar por CONTROLE
    grouped = df_stell.groupby("CONTROLE", as_index=False)
    faturados_set = set(df_faturados["CONTROLE"].astype(str))

    def determinar_lead_time(rotas):
        ### ab: checa betim/goiana e retorna lead_time e situ
        rotas_upper = [str(r).upper() for r in rotas]
        tem_goiana = any("GOIANA" in r for r in rotas_upper)
        tem_betim  = any("BETIM"  in r for r in rotas_upper)
        if tem_goiana:
            ### ab: prevalece goiana => 16 + 1
            return 16 + 1, "OK"
        elif tem_betim:
            ### ab: betim => 8 + 1
            return 8 + 1, "OK"
        else:
            ### ab: se nao achou rota => falta_rota
            return None, "FALTA_ROTA"

    def checar_lote_ida_ou_retorno(df_lote):
        ### ab: verifica se lote e so ida ou so retorno
        textos = df_lote["DTL_DESCRI"].astype(str).str.upper()
        encontrou_ida = textos.str.contains("IDA").any()
        encontrou_retorno = textos.str.contains("RETORNO").any()
        if encontrou_ida and encontrou_retorno:
            return "MISTO"
        if (not encontrou_ida) and (not encontrou_retorno):
            return "INDEFINIDO"
        return "IDA" if encontrou_ida else "RETORNO"

    hoje = date.today()

    ### ab: percorrer cada CONTROLE
    for controle_valor, group_df in grouped:
        ctrl_str = str(controle_valor)

        ### ab: se ja faturado
        if ctrl_str in faturados_set:
            verificar_data.append({
                "CONTROLE": ctrl_str,
                "TIPO": "Nº Ativação",
                "OBSERVACAO": "Este controle já foi faturado."
            })
            continue

        ### ab: rota => lead_time
        rotas_controle = group_df["ROTA"].unique()
        lead_time, situacao_rota = determinar_lead_time(rotas_controle)
        if situacao_rota == "FALTA_ROTA":
            rotas_str = " / ".join(str(r) for r in rotas_controle)
            verificar_data.append({
                "CONTROLE": ctrl_str,
                "TIPO": "Rota",
                "OBSERVACAO": f"Verificar rota: {rotas_str} para saber o lead-time."
            })
            continue

        ### ab: checa se ha tipo doc != CT-e
        tipos_doc_unicos = group_df["TIPO DOC"].astype(str).unique()
        has_non_cte = any(tdoc.lower() != "ct-e" for tdoc in tipos_doc_unicos)

        ### ab: data emissao max + lead_time
        data_emissao_max = group_df["DATA EMISSAO"].max().date()
        vencimento_grupo = data_emissao_max + timedelta(days=lead_time)

        ### ab: regra antiga
        lotes_unicos = group_df["LOTE"].unique()
        tem_2_lotes = (len(lotes_unicos) > 1)
        pode_faturar_por_data = (hoje >= vencimento_grupo)
        if has_non_cte:
            pode_faturar = pode_faturar_por_data
        else:
            pode_faturar = pode_faturar_por_data or tem_2_lotes

        if not pode_faturar:
            verificar_data.append({
                "CONTROLE": ctrl_str,
                "TIPO": "Não Faturar",
                "OBSERVACAO": (
                    "Ainda não está disponível para faturar. "
                    "Possível motivo: não passou do lead-time e/ou "
                    "há TIPO DOC != 'CT-e'."
                )
            })
            continue

        ### ab: logica extra de lotes
        cte_complementar_presente = any("ct-e complementar" in tdoc.lower() for tdoc in tipos_doc_unicos)
        num_lotes = len(lotes_unicos)

        ### ab: checar se dentro de cada lote nao ha mistura de ida e retorno
        lotes_tipos = []
        erro_lote = False
        for lote in lotes_unicos:
            df_lote = group_df[group_df["LOTE"] == lote]
            tipo_lote = checar_lote_ida_ou_retorno(df_lote)
            if tipo_lote in ("MISTO", "INDEFINIDO"):
                verificar_data.append({
                    "CONTROLE": ctrl_str,
                    "TIPO": "Erro Lote",
                    "OBSERVACAO": (
                        f"No lote {lote}, encontramos mistura de 'IDA' e 'RETORNO' "
                        "ou não encontramos nenhum deles. Verificar DTL_DESCRI."
                    )
                })
                erro_lote = True
            else:
                lotes_tipos.append(tipo_lote)
        if erro_lote:
            continue

        ### ab: validar qtd lotes e se sao ida/retorno
        if not cte_complementar_presente:
            ### ab: sem ct-e compl => exige 2 lotes
            if num_lotes != 2:
                verificar_data.append({
                    "CONTROLE": ctrl_str,
                    "TIPO": "Erro Lote",
                    "OBSERVACAO": (
                        f"Este CONTROLE tem {num_lotes} lote(s), mas NÃO há CT-e Complementar. "
                        "Deveria ter exatamente 2 lotes (IDA e RETORNO)."
                    )
                })
                continue
            tipos_encontrados = set(lotes_tipos)
            if tipos_encontrados != {"IDA", "RETORNO"}:
                verificar_data.append({
                    "CONTROLE": ctrl_str,
                    "TIPO": "Erro Lote",
                    "OBSERVACAO": (
                        f"Com 2 lotes, esperávamos um lote de IDA e um de RETORNO. "
                        f"Encontrados: {tipos_encontrados}"
                    )
                })
                continue
        else:
            ### ab: com ct-e compl
            if num_lotes == 2:
                tipos_encontrados = set(lotes_tipos)
                if tipos_encontrados != {"IDA", "RETORNO"}:
                    verificar_data.append({
                        "CONTROLE": ctrl_str,
                        "TIPO": "Erro Lote",
                        "OBSERVACAO": (
                            "Há CT-e Complementar, mas foram encontrados 2 lotes que não são IDA e RETORNO. "
                            f"Encontrados: {tipos_encontrados}"
                        )
                    })
                    continue

        ### ab: passou validacoes => a faturar
        for idx, row in group_df.iterrows():
            data_emissao_row = row["DATA EMISSAO"].date()
            dtl_descri_row = row.get("DTL_DESCRI", "")
            final_venc_date = get_final_due_date(data_emissao_row, lead_time, dtl_descri_row)
            data_venc_str = final_venc_date.strftime("%d/%m/%Y")
            a_faturar_data.append({
                "CGCPAGADOR": row.get("CNPJ", ""),
                "FILIALDOC": row.get("FILIAL", ""),
                "DOCUMENTO": row.get("NUM DOC", ""),
                "SERIE": row.get("SERIE", ""),
                "ID": ctrl_str,
                # "VENCIMENTODOC": data_venc_str,
                "FRETETOTAL": row.get("VALOR", 0)
            })

    ### ab: salvar dataframes
    df_faturar = pd.DataFrame(a_faturar_data)
    df_verificar = pd.DataFrame(verificar_data)
    data_hoje = datetime.now().strftime("%d.%m.%y")
    nome_arquivo_saida = f"{data_hoje} - Stellantis.xlsx"
    caminho_saida = os.path.join(pasta_base, nome_arquivo_saida)
    with pd.ExcelWriter(caminho_saida, engine="openpyxl") as writer:
        df_faturar.to_excel(writer, sheet_name="A Faturar", index=False)
        df_verificar.to_excel(writer, sheet_name="Verificar", index=False)
    print(f"relatório gerado com sucesso em: {caminho_saida}")

    ### ab: csv saida
    csv_folder = os.path.join(pasta_base, "00. CSV")
    os.makedirs(csv_folder, exist_ok=True)
    nome_arquivo_csv = f"{data_hoje} - Stellantis.csv"
    caminho_csv = os.path.join(csv_folder, nome_arquivo_csv)
    df_faturar.to_csv(caminho_csv, index=False, sep=";")
    print(f"arquivo csv gerado com sucesso em: {caminho_csv}")

if __name__ == "__main__":
    main()
