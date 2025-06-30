import gradio as gr
import pandas as pd
import numpy as np
import os
from datetime import datetime, timedelta
from geopy.distance import geodesic

def traduzir_dia_semana(date_obj):
    dias_pt = {
        "Monday": "segunda-feira",
        "Tuesday": "terça-feira",
        "Wednesday": "quarta-feira",
        "Thursday": "quinta-feira",
        "Friday": "sexta-feira",
        "Saturday": "sábado",
        "Sunday": "domingo"
    }
    return dias_pt[date_obj.strftime('%A')]

def formatar_nome_simples(nome):
    nome = nome.strip()
    nome = nome.replace("CI ", "").replace("Ci ", "").replace("C i ", "").replace("C I ", "")
    partes = nome.split()
    if partes[0].lower() in ['ana', 'maria'] and len(partes) > 1:
        return " ".join(partes[:2])
    else:
        return partes[0]

def gerar_mensagem_personalizada(
    nome_profissional, nome_cliente, data_servico, servico,
    duracao, rua, numero, complemento, bairro, cidade, latitude, longitude,
    ja_atendeu
):
    nome_profissional_fmt = formatar_nome_simples(nome_profissional)
    nome_cliente_fmt = nome_cliente.split()[0].strip().title()
    data_dt = pd.to_datetime(data_servico, errors="coerce")
    if pd.isnull(data_dt):
        data_formatada = ""
        dia_semana = ""
    else:
        dia_semana = traduzir_dia_semana(data_dt)
        data_formatada = data_dt.strftime("%d/%m/%Y")
    data_linha = f"{dia_semana}, {data_formatada}"
    endereco_str = f"{rua}, {numero}"
    if complemento and str(complemento).strip().lower() not in ["nan", "none", "-"]:
        endereco_str += f", {complemento}"
    if pd.notnull(latitude) and pd.notnull(longitude):
        maps_url = f"https://maps.google.com/?q={latitude},{longitude}"
    else:
        maps_url = ""
    fechamento = (
        "SIM ou não para o aceite" if ja_atendeu
        else "responda com SIM caso tenha disponibilidade!"
    )
    rodape = (
        "O atendimento será confirmado após o recebimento das informações completas do atendimento, Nome e observações do cliente. Ok?\n\n"
        "Lembre que o cliente irá receber o profissional indicado pela Vavivê, então lembre-se das nossas 3 confirmações do atendimento!\n\n"
        "CONFIRME SE O ATENDINEMTO AINDA ESTÁ VÁLIDO\n\n"
        "Abs, Vavivê!"
    )
    mensagem = f"""Olá {nome_profissional_fmt}!

Temos uma oportunidade especial para você nesta região! Quer assumir essa demanda? Está dentro da sua rota!

**Cliente:** {nome_cliente_fmt}
📅 **Data:** {data_linha}
🛠️ **Serviço:** {servico}
⏱️ **Duração do Atendimento:** {duracao}
📍 **Endereço:** {endereco_str}
📍 **Bairro:** {bairro}
🏙️ **Cidade:** {cidade}
{"🌎 [Abrir no Google Maps](" + maps_url + ")" if maps_url else ""}

{fechamento}

{rodape}
"""
    return mensagem

def padronizar_cpf_cnpj(coluna):
    return (
        coluna.astype(str)
        .str.replace(r'\D', '', regex=True)
        .str.zfill(11)  # Se só CPF, use 11; se também CNPJ, use 14
        .str.strip()
    )

def processar_excel(arquivo):
    # Nome do arquivo de saída
    final_path = "matriz_rotas_v1_com_criterios.xlsx"
    output_dir = "output_colab"
    os.makedirs(output_dir, exist_ok=True)
    file_path = arquivo.name

    # ==================================
    # 2️⃣ Tratar aba Clientes
    # ==================================
    df_clientes_raw = pd.read_excel(file_path, sheet_name="Clientes")
    df_clientes = df_clientes_raw[[
        "ID","UpdatedAt","celular","cpf",
        "endereco-1-bairro","endereco-1-cidade","endereco-1-complemento",
        "endereco-1-estado","endereco-1-latitude","endereco-1-longitude",
        "endereco-1-numero","endereco-1-rua","nome"
    ]].copy()
    df_clientes["ID Cliente"] = (
        df_clientes["ID"].fillna("0").astype(str)
        .str.replace(r"\.0$", "", regex=True).str.strip()
    )
    df_clientes["CPF_CNPJ"] = padronizar_cpf_cnpj(df_clientes["cpf"])
    df_clientes["Celular"] = df_clientes["celular"].astype(str).str.strip()
    df_clientes["Complemento"] = df_clientes["endereco-1-complemento"].astype(str).str.strip()
    df_clientes["Número"] = df_clientes["endereco-1-numero"].astype(str).str.strip()
    df_clientes["Nome Cliente"] = df_clientes["nome"].astype(str).str.strip()
    df_clientes = df_clientes.rename(columns={
        "endereco-1-bairro": "Bairro",
        "endereco-1-cidade": "Cidade",
        "endereco-1-estado": "Estado",
        "endereco-1-latitude": "Latitude Cliente",
        "endereco-1-longitude": "Longitude Cliente",
        "endereco-1-rua": "Rua"
    })
    df_clientes["Latitude Cliente"] = pd.to_numeric(df_clientes["Latitude Cliente"], errors="coerce")
    df_clientes["Longitude Cliente"] = pd.to_numeric(df_clientes["Longitude Cliente"], errors="coerce")
    coord_invertida = df_clientes["Latitude Cliente"] < -40
    if coord_invertida.any():
        lat_temp = df_clientes.loc[coord_invertida, "Latitude Cliente"].copy()
        df_clientes.loc[coord_invertida, "Latitude Cliente"] = df_clientes.loc[coord_invertida, "Longitude Cliente"]
        df_clientes.loc[coord_invertida, "Longitude Cliente"] = lat_temp
    df_clientes["coordenadas_validas"] = df_clientes["Latitude Cliente"].notnull() & df_clientes["Longitude Cliente"].notnull()
    df_clientes = df_clientes.sort_values(by=["CPF_CNPJ", "coordenadas_validas"], ascending=[True, False])
    df_clientes = df_clientes.drop_duplicates(subset="CPF_CNPJ", keep="first")
    df_clientes.drop(columns=["coordenadas_validas"], inplace=True)
    df_clientes = df_clientes[[
        "ID Cliente","UpdatedAt","Celular","CPF_CNPJ",
        "Bairro","Cidade","Complemento","Estado","Latitude Cliente","Longitude Cliente",
        "Número","Rua","Nome Cliente"
    ]]

    # ==================================
    # 3️⃣ Tratar aba Profissionais
    # ==================================
    df_profissionais_raw = pd.read_excel(file_path, sheet_name="Profissionais")
    df_profissionais = df_profissionais_raw[[
        "ID","atendimentos_feitos","celular","cpf",
        "endereco-bairro","endereco-cidade","endereco-complemento","endereco-estado",
        "endereco-latitude","endereco-longitude","endereco-numero","endereco-rua","nome"
    ]].copy()
    df_profissionais["ID Prestador"] = (
        df_profissionais["ID"].fillna("0").astype(str)
        .str.replace(r"\.0$", "", regex=True).str.strip()
    )
    df_profissionais["Qtd Atendimentos"] = df_profissionais["atendimentos_feitos"].fillna(0).astype(int)
    df_profissionais["Celular"] = df_profissionais["celular"].astype(str).str.strip()
    df_profissionais["cpf"] = (
        df_profissionais["cpf"].astype(str).str.replace(r"\D", "", regex=True).str.strip()
    )
    df_profissionais["Complemento"] = df_profissionais["endereco-complemento"].astype(str).str.strip()
    df_profissionais["Número"] = df_profissionais["endereco-numero"].astype(str).str.strip()
    df_profissionais["Nome Prestador"] = df_profissionais["nome"].astype(str).str.strip()
    df_profissionais = df_profissionais.rename(columns={
        "endereco-bairro": "Bairro",
        "endereco-cidade": "Cidade",
        "endereco-estado": "Estado",
        "endereco-latitude": "Latitude Profissional",
        "endereco-longitude": "Longitude Profissional",
        "endereco-rua": "Rua"
    })
    df_profissionais = df_profissionais[~df_profissionais["Nome Prestador"].str.contains("inativo", case=False, na=False)].copy()
    df_profissionais["Latitude Profissional"] = pd.to_numeric(df_profissionais["Latitude Profissional"], errors="coerce")
    df_profissionais["Longitude Profissional"] = pd.to_numeric(df_profissionais["Longitude Profissional"], errors="coerce")
    df_profissionais = df_profissionais[
        df_profissionais["Latitude Profissional"].notnull() &
        df_profissionais["Longitude Profissional"].notnull()
    ].copy()
    df_profissionais = df_profissionais[[
        "ID Prestador","Qtd Atendimentos","Celular","cpf",
        "Bairro","Cidade","Complemento","Estado","Latitude Profissional","Longitude Profissional",
        "Número","Rua","Nome Prestador"
    ]]

    # ==================================
    # 4️⃣ Preferências, Bloqueio, Queridinhos, Sumidinhos
    # ==================================
    df_preferencias_raw = pd.read_excel(file_path, sheet_name="Preferencias")
    df_preferencias = df_preferencias_raw[[
        "CPF/CNPJ","Cliente","ID Profissional","Prestador"
    ]].copy()
    df_preferencias["CPF_CNPJ"] = padronizar_cpf_cnpj(df_preferencias["CPF/CNPJ"])
    df_preferencias["Nome Cliente"] = df_preferencias["Cliente"].astype(str).str.strip()
    df_preferencias["ID Prestador"] = (
        df_preferencias["ID Profissional"].fillna("0").astype(str)
        .str.replace(r"\.0$", "", regex=True).str.strip()
    )
    df_preferencias["Nome Prestador"] = df_preferencias["Prestador"].astype(str).str.strip()
    df_preferencias = df_preferencias[[
        "CPF_CNPJ","Nome Cliente","ID Prestador","Nome Prestador"
    ]]

    df_bloqueio_raw = pd.read_excel(file_path, sheet_name="Bloqueio")
    df_bloqueio = df_bloqueio_raw[[
        "CPF/CNPJ","Cliente","ID Profissional","Prestador"
    ]].copy()
    df_bloqueio["CPF_CNPJ"] = padronizar_cpf_cnpj(df_bloqueio["CPF/CNPJ"])
    df_bloqueio["Nome Cliente"] = df_bloqueio["Cliente"].astype(str).str.strip()
    df_bloqueio["ID Prestador"] = (
        df_bloqueio["ID Profissional"].fillna("0").astype(str)
        .str.replace(r"\.0$", "", regex=True).str.strip()
    )
    df_bloqueio["Nome Prestador"] = df_bloqueio["Prestador"].astype(str).str.strip()
    df_bloqueio = df_bloqueio[[
        "CPF_CNPJ","Nome Cliente","ID Prestador","Nome Prestador"
    ]]

    df_queridinhos_raw = pd.read_excel(file_path, sheet_name="Profissionais Preferenciais")
    df_queridinhos = df_queridinhos_raw[[
        "ID Profissional","Profissional"
    ]].copy()
    df_queridinhos["ID Prestador"] = (
        df_queridinhos["ID Profissional"].fillna("0").astype(str)
        .str.replace(r"\.0$", "", regex=True).str.strip()
    )
    df_queridinhos["Nome Prestador"] = df_queridinhos["Profissional"].astype(str).str.strip()
    df_queridinhos = df_queridinhos[["ID Prestador","Nome Prestador"]]

    df_sumidinhos_raw = pd.read_excel(file_path, sheet_name="Baixa Disponibilidade")
    df_sumidinhos = df_sumidinhos_raw[[
        "ID Profissional","Profissional"
    ]].copy()
    df_sumidinhos["ID Prestador"] = (
        df_sumidinhos["ID Profissional"].fillna("0").astype(str)
        .str.replace(r"\.0$", "", regex=True).str.strip()
    )
    df_sumidinhos["Nome Prestador"] = df_sumidinhos["Profissional"].astype(str).str.strip()
    df_sumidinhos = df_sumidinhos[["ID Prestador","Nome Prestador"]]

    # ==================================
    # 5️⃣ Atendimentos e Histórico 60 dias
    # ==================================
    df_atendimentos = pd.read_excel(file_path, sheet_name="Atendimentos")
    colunas_desejadas = [
        "OS","Status Serviço","Data 1","Plano","CPF/ CNPJ","Cliente","Serviço",
        "Horas de serviço","Hora de entrada","Observações atendimento",
        "Observações prestador","Ponto de Referencia","#Num Prestador","Prestador"
    ]
    df_atendimentos = df_atendimentos[colunas_desejadas].copy()
    df_atendimentos["Data 1"] = pd.to_datetime(df_atendimentos["Data 1"], errors="coerce")
    df_atendimentos["CPF_CNPJ"] = padronizar_cpf_cnpj(df_atendimentos["CPF/ CNPJ"])
    df_atendimentos["Cliente"] = df_atendimentos["Cliente"].astype(str).str.strip()
    df_atendimentos["Duração do Serviço"] = df_atendimentos["Horas de serviço"]
    df_atendimentos["ID Prestador"] = (
        df_atendimentos["#Num Prestador"].fillna("0").astype(str)
        .str.replace(r"\.0$", "", regex=True).str.strip()
    )

    hoje = datetime.now().date()
    limite = hoje - timedelta(days=60)
    data1_datetime = pd.to_datetime(df_atendimentos["Data 1"], errors="coerce")
    df_historico_60_dias = df_atendimentos[
        (df_atendimentos["Status Serviço"].str.lower() != "cancelado") &
        (data1_datetime.dt.date < hoje) &
        (data1_datetime.dt.date >= limite)
    ].copy()
    df_historico_60_dias = df_historico_60_dias[[
        "CPF_CNPJ","Cliente","Data 1","Status Serviço","Serviço",
        "Duração do Serviço","Hora de entrada","ID Prestador","Prestador"
    ]]

    # Cliente x Prestador histórico
    df_cliente_prestador = df_historico_60_dias.groupby(
        ["CPF_CNPJ","ID Prestador"]
    ).size().reset_index(name="Qtd Atendimentos Cliente-Prestador")

    # Qtd atendimentos por prestador histórico
    df_qtd_por_prestador = df_historico_60_dias.groupby(
        "ID Prestador"
    ).size().reset_index(name="Qtd Atendimentos Prestador")

    # ==================================
    # 6️⃣ Calcular distâncias
    # ==================================
    df_clientes_coord = df_clientes[["CPF_CNPJ","Latitude Cliente","Longitude Cliente"]].dropna().drop_duplicates("CPF_CNPJ")
    df_profissionais_coord = df_profissionais[["ID Prestador","Latitude Profissional","Longitude Profissional"]].dropna().drop_duplicates("ID Prestador")

    distancias = []
    for _, cliente in df_clientes_coord.iterrows():
        coord_cliente = (cliente["Latitude Cliente"], cliente["Longitude Cliente"])
        for _, profissional in df_profissionais_coord.iterrows():
            coord_prof = (profissional["Latitude Profissional"], profissional["Longitude Profissional"])
            distancia_km = round(geodesic(coord_cliente, coord_prof).km, 2)
            distancias.append({
                "CPF_CNPJ": cliente["CPF_CNPJ"],
                "ID Prestador": profissional["ID Prestador"],
                "Distância (km)": distancia_km
            })
    df_distancias = pd.DataFrame(distancias)
    df_distancias_alerta = df_distancias[df_distancias["Distância (km)"] > 1000]

    # ==================================
    # 7️⃣ Preferencias e bloqueios com join geográfico
    # ==================================
    df_preferencias_completo = df_preferencias.merge(
        df_clientes_coord, on="CPF_CNPJ", how="left"
    ).merge(
        df_profissionais_coord, on="ID Prestador", how="left"
    )
    df_preferencias_completo = df_preferencias_completo[[
        "CPF_CNPJ","Nome Cliente","ID Prestador","Nome Prestador",
        "Latitude Cliente","Longitude Cliente",
        "Latitude Profissional","Longitude Profissional"
    ]]

    df_bloqueio_completo = df_bloqueio.merge(
        df_clientes_coord, on="CPF_CNPJ", how="left"
    ).merge(
        df_profissionais_coord, on="ID Prestador", how="left"
    )
    df_bloqueio_completo = df_bloqueio_completo[[
        "CPF_CNPJ","Nome Cliente","ID Prestador","Nome Prestador",
        "Latitude Cliente","Longitude Cliente",
        "Latitude Profissional","Longitude Profissional"
    ]]

    # ==================================
    # 8️⃣ Atendimento futuro com e sem localização
    # ==================================
    ontem = datetime.now().date() - timedelta(days=1)
    df_futuros = df_atendimentos[
        (df_atendimentos["Status Serviço"].str.lower() != "cancelado") &
        (df_atendimentos["Data 1"].dt.date > ontem)
    ].copy()
    df_futuros_com_clientes = df_futuros.merge(
        df_clientes_coord, on="CPF_CNPJ", how="left"
    )
    colunas_uteis = [
        "OS","Data 1","Status Serviço","CPF_CNPJ","Cliente","Serviço",
        "Duração do Serviço","Hora de entrada","Ponto de Referencia",
        "ID Prestador","Prestador","Latitude Cliente","Longitude Cliente","Plano"
    ]
    df_atendimentos_futuros_validos = df_futuros_com_clientes[
        df_futuros_com_clientes["Latitude Cliente"].notnull() &
        df_futuros_com_clientes["Longitude Cliente"].notnull()
    ][colunas_uteis].copy()

    df_atendimentos_sem_localizacao = df_futuros_com_clientes[
        df_futuros_com_clientes["Latitude Cliente"].isnull() |
        df_futuros_com_clientes["Longitude Cliente"].isnull()
    ][colunas_uteis].copy()

    # ==================================
    # MATRIZ DE ROTAS
    # ==================================
    matriz_resultado_corrigida = []

    for _, atendimento in df_atendimentos_futuros_validos.iterrows():
        os_id = atendimento["OS"]
        cpf = atendimento["CPF_CNPJ"]
        nome_cliente = atendimento["Cliente"]
        data_1 = atendimento["Data 1"].strftime("%d/%m/%Y")
        servico = atendimento["Serviço"]
        duracao_servico = atendimento["Duração do Serviço"]
        hora_entrada = atendimento["Hora de entrada"]
        ponto_referencia = atendimento["Ponto de Referencia"]
        lat_cliente = atendimento["Latitude Cliente"]
        lon_cliente = atendimento["Longitude Cliente"]
        plano = atendimento.get("Plano", "")

        # Bloqueados (sempre string)
        bloqueados = (
            df_bloqueio_completo[df_bloqueio_completo["CPF_CNPJ"] == cpf]["ID Prestador"]
            .astype(str).str.strip().tolist()
        )

        # Candidatos SEM bloqueados
        candidatos = df_profissionais[
            ~df_profissionais["ID Prestador"].astype(str).str.strip().isin(bloqueados)
        ].copy()

        linha = {
            "OS": os_id,
            "CPF_CNPJ": cpf,
            "Nome Cliente": nome_cliente,
            "Plano": plano,
            "Data 1": data_1,
            "Serviço": servico,
            "Duração do Serviço": duracao_servico,
            "Hora de entrada": hora_entrada,
            "Ponto de Referencia": ponto_referencia
        }

        cliente_match = df_clientes[df_clientes["CPF_CNPJ"] == cpf]
        cliente_info = cliente_match.iloc[0] if not cliente_match.empty else None

        if cliente_info is not None:
            rua = cliente_info["Rua"]
            numero = cliente_info["Número"]
            complemento = cliente_info["Complemento"]
            bairro = cliente_info["Bairro"]
            cidade = cliente_info["Cidade"]
            latitude = cliente_info["Latitude Cliente"]
            longitude = cliente_info["Longitude Cliente"]
        else:
            rua = numero = complemento = bairro = cidade = latitude = longitude = ""

        # 1️⃣ Sempre verifica e aplica preferida PRIMEIRO (não pode nunca ser “tomada” pelo histórico!)
        preferencia_cliente_df = df_preferencias_completo[df_preferencias_completo["CPF_CNPJ"] == cpf]
        preferida_id = None
        preferida_row_data = None

        if not preferencia_cliente_df.empty:
            id_preferida_temp = str(preferencia_cliente_df.iloc[0]["ID Prestador"]).strip()
            if id_preferida_temp not in bloqueados:
                profissional_preferida_info = df_profissionais[
                    df_profissionais["ID Prestador"].astype(str).str.strip() == id_preferida_temp
                ]
                if not profissional_preferida_info.empty:
                    lat_prof_pref = profissional_preferida_info.iloc[0]["Latitude Profissional"]
                    lon_prof_pref = profissional_preferida_info.iloc[0]["Longitude Profissional"]
                    if pd.notnull(lat_prof_pref) and pd.notnull(lon_prof_pref):
                        preferida_id = id_preferida_temp
                        qtd_atend_cliente_pref = df_cliente_prestador[
                            (df_cliente_prestador["CPF_CNPJ"] == cpf) &
                            (df_cliente_prestador["ID Prestador"] == preferida_id)
                        ]["Qtd Atendimentos Cliente-Prestador"]
                        qtd_atend_cliente_pref = qtd_atend_cliente_pref.iloc[0] if not qtd_atend_cliente_pref.empty else 1
                        qtd_atend_total_pref = df_qtd_por_prestador[
                            df_qtd_por_prestador["ID Prestador"] == preferida_id
                        ]["Qtd Atendimentos Prestador"]
                        qtd_atend_total_pref = qtd_atend_total_pref.iloc[0] if not qtd_atend_total_pref.empty else 0
                        distancia_pref_df = df_distancias[
                            (df_distancias["CPF_CNPJ"] == cpf) &
                            (df_distancias["ID Prestador"] == preferida_id)
                        ]
                        distancia_pref = distancia_pref_df["Distância (km)"].iloc[0] if not distancia_pref_df.empty else 9999
                        preferida_criterio = f"cliente: {qtd_atend_cliente_pref} | total: {qtd_atend_total_pref} — {distancia_pref:.2f} km"
                        preferida_row_data = {
                            "Classificação da Profissional 1": 1,
                            "Critério 1": preferida_criterio,
                            "Nome Prestador 1": profissional_preferida_info.iloc[0]["Nome Prestador"],
                            "Celular 1": profissional_preferida_info.iloc[0]["Celular"],
                            "Mensagem 1": gerar_mensagem_personalizada(
                                profissional_preferida_info.iloc[0]["Nome Prestador"],
                                nome_cliente, data_1, servico,
                                duracao_servico, rua, numero, complemento, bairro, cidade,
                                latitude, longitude, ja_atendeu=True
                            ),
                            "Critério Utilizado 1": "Preferência do Cliente"
                        }
                        candidatos = candidatos[candidatos["ID Prestador"].astype(str).str.strip() != preferida_id].copy()
                        linha.update(preferida_row_data)
                        start_rank = 2
                    else:
                        start_rank = 1
                else:
                    start_rank = 1
            else:
                start_rank = 1
        else:
            start_rank = 1

        # 2️⃣ Agora faz o merge do histórico para o restante dos candidatos
        candidatos_merged = candidatos.merge(
            df_cliente_prestador[df_cliente_prestador["CPF_CNPJ"] == cpf][["ID Prestador", "Qtd Atendimentos Cliente-Prestador"]],
            on="ID Prestador", how="left"
        ).merge(
            df_qtd_por_prestador, on="ID Prestador", how="left"
        ).merge(
            df_distancias[df_distancias["CPF_CNPJ"] == cpf][["ID Prestador", "Distância (km)"]],
            on="ID Prestador", how="left"
        )
        candidatos_merged["Qtd Atendimentos Cliente-Prestador"] = candidatos_merged["Qtd Atendimentos Cliente-Prestador"].fillna(0).astype(int)
        candidatos_merged["Qtd Atendimentos Prestador"] = candidatos_merged["Qtd Atendimentos Prestador"].fillna(0).astype(int)
        candidatos_merged = candidatos_merged.sort_values(
            by=["Qtd Atendimentos Cliente-Prestador", "Distância (km)"],
            ascending=[False, True]
        )

        if start_rank == 1:
            if candidatos_merged["Qtd Atendimentos Cliente-Prestador"].max() > 0:
                max_atend = candidatos_merged["Qtd Atendimentos Cliente-Prestador"].max()
                top_candidates = candidatos_merged[candidatos_merged["Qtd Atendimentos Cliente-Prestador"] == max_atend]
                top1 = top_candidates.sort_values("Distância (km)").iloc[0]
                linha["Classificação da Profissional 1"] = 1
                linha["Critério 1"] = f"cliente: {top1['Qtd Atendimentos Cliente-Prestador']} | total: {top1['Qtd Atendimentos Prestador']} — {top1['Distância (km)']:.2f} km"
                linha["Nome Prestador 1"] = top1["Nome Prestador"]
                linha["Celular 1"] = top1["Celular"]
                linha["Mensagem 1"] = gerar_mensagem_personalizada(
                    top1["Nome Prestador"], nome_cliente, data_1, servico,
                    duracao_servico, rua, numero, complemento, bairro, cidade,
                    latitude, longitude, ja_atendeu=True
                )
                linha["Critério Utilizado 1"] = "Mais atendeu o cliente"
                candidatos_merged = candidatos_merged[candidatos_merged["ID Prestador"] != top1["ID Prestador"]].copy()
                start_rank = 2

        for i, (_, row) in enumerate(candidatos_merged.head(20 - (start_rank - 1)).iterrows(), start=start_rank):
            crit = f"cliente: {row.get('Qtd Atendimentos Cliente-Prestador',0)} | total: {row.get('Qtd Atendimentos Prestador',0)} — {row.get('Distância (km)',0):.2f} km"
            linha[f"Classificação da Profissional {i}"] = i
            linha[f"Critério {i}"] = crit
            linha[f"Nome Prestador {i}"] = row["Nome Prestador"]
            linha[f"Celular {i}"] = row["Celular"]
            linha[f"Mensagem {i}"] = gerar_mensagem_personalizada(
                row["Nome Prestador"], nome_cliente, data_1, servico,
                duracao_servico, rua, numero, complemento, bairro, cidade,
                latitude, longitude, ja_atendeu=(row.get('Qtd Atendimentos Cliente-Prestador',0)>0)
            )
            linha[f"Critério Utilizado {i}"] = (
                "Mais atendeu o cliente" if row.get('Qtd Atendimentos Cliente-Prestador',0)>0 else "Mais próxima geograficamente"
            )

        matriz_resultado_corrigida.append(linha)

    df_matriz_rotas = pd.DataFrame(matriz_resultado_corrigida)

    # Preencher colunas de Profissionais/Critérios que não foram preenchidas (menos de 20 candidatos)
    for i in range(1, 21):
        if f"Classificação da Profissional {i}" not in df_matriz_rotas.columns:
            df_matriz_rotas[f"Classificação da Profissional {i}"] = pd.NA
        if f"Critério {i}" not in df_matriz_rotas.columns:
            df_matriz_rotas[f"Critério {i}"] = pd.NA
        if f"Nome Prestador {i}" not in df_matriz_rotas.columns:
            df_matriz_rotas[f"Nome Prestador {i}"] = pd.NA
        if f"Celular {i}" not in df_matriz_rotas.columns:
            df_matriz_rotas[f"Celular {i}"] = pd.NA
        if f"Mensagem {i}" not in df_matriz_rotas.columns:
            df_matriz_rotas[f"Mensagem {i}"] = pd.NA
        if f"Critério Utilizado {i}" not in df_matriz_rotas.columns:
            df_matriz_rotas[f"Critério Utilizado {i}"] = pd.NA

    base_cols = ["OS", "CPF_CNPJ", "Nome Cliente", "Data 1", "Serviço", "Duração do Serviço", "Hora de entrada", "Ponto de Referencia"]
    prestador_cols = []
    for i in range(1, 21):
        prestador_cols.extend([
            f"Classificação da Profissional {i}",
            f"Critério {i}",
            f"Nome Prestador {i}",
            f"Celular {i}",
            f"Mensagem {i}",
            f"Critério Utilizado {i}"
        ])
    df_matriz_rotas = df_matriz_rotas[base_cols + prestador_cols]

    # Salva Excel final para download
    file_path_final = os.path.join(output_dir, final_path)
    df_matriz_rotas.to_excel(file_path_final, index=False)
    return file_path_final

iface = gr.Interface(
    fn=processar_excel,
    inputs=gr.File(label="Envie seu arquivo Excel (original)"),
    outputs=gr.File(label="Baixe o arquivo tratado"),
    title="Portal Vavivê - Processamento de Excel de Rotas",
    description="Faça upload do seu arquivo Excel original e baixe o resultado tratado e consolidado.",
    allow_flagging="never"
)

if __name__ == "__main__":
    iface.launch()
