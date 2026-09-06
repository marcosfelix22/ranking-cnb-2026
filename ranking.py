import os
import requests
import pandas as pd

# CONFIGURAÇÕES
CLIENT_ID = (os.environ.get('CLIENT_ID') or '').strip()
CLIENT_SECRET = (os.environ.get('CLIENT_SECRET') or '').strip()
SUPABASE_URL = (os.environ.get('SUPABASE_URL') or '').strip()
SUPABASE_KEY = (os.environ.get('SUPABASE_KEY') or '').strip()
NOME_ARQUIVO = 'Ranking_CNB_2026.xlsx'

def buscar_atletas_supabase():
    """Busca a lista de atletas e seus refresh_tokens salvos no Supabase pelo Lovable"""
    headers = {
        'apikey': SUPABASE_KEY,
        'Authorization': f'Bearer {SUPABASE_KEY}'
    }
    # Consulta a tabela 'atletas' do Supabase
    url = f"{SUPABASE_URL}/rest/v1/atletas?select=nome,refresh_token"
    res = requests.get(url, headers=headers)
    if res.status_code == 200:
        return res.json()
    else:
        print(f"Erro ao buscar atletas no Supabase: {res.status_code} - {res.text}")
        return []

def obter_access_token(refresh_token):
    payload = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'refresh_token': refresh_token.strip(),
        'grant_type': 'refresh_token'
    }
    res = requests.post("https://www.strava.com/oauth/token", data=payload)
    return res.json().get('access_token') if res.status_code == 200 else None

def formatar_km(valor):
    return f"{valor:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".") + " km"

def formatar_alt(valor):
    return f"{int(valor):,}".replace(",", ".") + " m"

print("--- EXECUTANDO ATUALIZAÇÃO AUTOMÁTICA DO RANKING CNB 2026 ---")

# 1. Puxa todos os atletas cadastrados pelo Lovable
lista_atletas = buscar_atletas_supabase()
dados_ranking = []

for atleta in lista_atletas:
    nome = atleta.get('nome')
    ref_token = atleta.get('refresh_token')
    
    if not ref_token:
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
            if act.get('type') in ['Run', 'TrailRun'] and act.get('start_date', '').startswith('2026'):
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

# 2. Ordena e gera a planilha
if dados_ranking:
    df = pd.DataFrame(dados_ranking)
    df = df.sort_values(by='KM Total Bruto', ascending=False).drop(columns=['KM Total Bruto'])
else:
    df = pd.DataFrame(columns=['Atleta', 'KM Total', 'Altimetria (m)', 'Treinos'])

with pd.ExcelWriter(NOME_ARQUIVO, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Ranking CNB 2026', index=False)

print(f"✅ Sincronização da planilha {NOME_ARQUIVO} concluída com sucesso!")
