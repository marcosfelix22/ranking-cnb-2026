import os
import requests
import pandas as pd

# --- CONFIGURAÇÕES ---
CLIENT_ID = (os.environ.get('CLIENT_ID') or '').strip()
CLIENT_SECRET = (os.environ.get('CLIENT_SECRET') or '').strip()

# Endpoint da API criada no Lovable Cloud
URL_API_LOVABLE = "https://desafiocnb.lovable.app/api/atletas"
NOME_ARQUIVO = 'Ranking_CNB_2026.xlsx'

def buscar_atletas_lovable():
    """Busca a lista de atletas cadastrados via Lovable Cloud"""
    try:
        res = requests.get(URL_API_LOVABLE, timeout=15)
        if res.status_code == 200:
            return res.json()
        else:
            print(f"❌ Erro ao buscar atletas na API ({res.status_code}): {res.text}")
            return []
    except Exception as e:
        print(f"❌ Exceção ao conectar na API do Lovable: {e}")
        return []

def obter_access_token(refresh_token):
    if not refresh_token:
        return None

    payload = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'refresh_token': refresh_token.strip(),
        'grant_type': 'refresh_token'
    }
    
    try:
        res = requests.post("https://www.strava.com/oauth/token", data=payload)
        return res.json().get('access_token') if res.status_code == 200 else None
    except:
        return None

def formatar_km(valor):
    return f"{valor:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".") + " km"

def formatar_alt(valor):
    return f"{int(valor):,}".replace(",", ".") + " m"

print("--- EXECUTANDO ATUALIZAÇÃO AUTOMÁTICA DO RANKING CNB 2026 ---")

lista_atletas = buscar_atletas_lovable()
dados_ranking = []

if not lista_atletas:
    print("⚠️ Nenhum atleta retornado pela API ou tabela vazia.")

for atleta in lista_atletas:
    nome = atleta.get('nome') or atleta.get('name') or 'Atleta Sem Nome'
    ref_token = atleta.get('refresh_token')
    
    if not ref_token:
        print(f"Aviso: Atleta '{nome}' sem refresh_token salvo.")
        continue

    access_token = obter_access_token(ref_token)
    if not access_token:
        print(f"Erro ao obter access_token para '{nome}'.")
        continue

    headers = {'Authorization': f'Bearer {access_token}'}
    url = "https://www.strava.com/api/v3/athlete/activities"
    res = requests.get(url, headers=headers, params={'per_page': 200, 'page': 1})
    
    if res.status_code == 200:
        atividades = res.json()
        km_total, alt_total, treinos = 0.0, 0.0, 0
        
        for act in atividades:
            tipo = act.get('type')
            data_inicio = act.get('start_date', '')
            
            if tipo in ['Run', 'TrailRun'] and data_inicio.startswith('2026'):
                km_total += act.get('distance', 0.0) / 1000.0
                alt_total += act.get('total_elevation_gain', 0.0)
                treinos += 1

        dados_ranking.append({
            'Atleta': nome,
            'KM Total Bruto': km_total,
            'KM Total': formatar_km(km_total),
            'Altimetria (m)': formatar_alt(alt_total),
            'Treinos': treinos
        })
        print(f"✓ {nome}: {formatar_km(km_total)} em {treinos} treinos")
    else:
        print(f"Erro na consulta do Strava para {nome}: Status {res.status_code}")

if dados_ranking:
    df = pd.DataFrame(dados_ranking)
    df = df.sort_values(by='KM Total Bruto', ascending=False).drop(columns=['KM Total Bruto'])
else:
    df = pd.DataFrame(columns=['Atleta', 'KM Total', 'Altimetria (m)', 'Treinos'])

with pd.ExcelWriter(NOME_ARQUIVO, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Ranking CNB 2026', index=False)

print(f"✅ Sincronização da planilha {NOME_ARQUIVO} concluída com sucesso!")
