import os
import requests
import pandas as pd

# --- CONFIGURAÇÕES ---
CLIENT_ID = (os.environ.get('CLIENT_ID') or '').strip()
CLIENT_SECRET = (os.environ.get('CLIENT_SECRET') or '').strip()
NOME_ARQUIVO = 'Ranking_CNB_2026.xlsx'

# Lista única de Atletas do Desafio "Ranking CNB 2026"
ATLETAS = {
    "Marcos Felix": os.environ.get('TOKEN_MARCOS'),
    # Novos atletas entram aqui conforme enviarem autorização:
    # "Flávio Brayner": os.environ.get('TOKEN_FLAVIO'),
}

def obter_access_token(refresh_token):
    if not refresh_token:
        return None
        
    ref_token_limpo = refresh_token.strip()
    payload = {
        'client_id': CLIENT_ID,
        'client_secret': CLIENT_SECRET,
        'refresh_token': ref_token_limpo,
        'grant_type': 'refresh_token'
    }
    
    try:
        res = requests.post("https://www.strava.com/oauth/token", data=payload)
        if res.status_code == 200:
            return res.json().get('access_token')
        else:
            print(f"Erro ao renovar token ({res.status_code}): {res.text}")
            return None
    except Exception as e:
        print(f"Exceção ao obter access_token: {e}")
        return None

def formatar_km(valor):
    return f"{valor:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".") + " km"

def formatar_alt(valor):
    return f"{int(valor):,}".replace(",", ".") + " m"

dados_ranking = []

for nome_atleta, ref_token in ATLETAS.items():
    if not ref_token:
        print(f"Aviso: Secret do atleta '{nome_atleta}' não configurado.")
        continue

    access_token = obter_access_token(ref_token)
    if not access_token:
        print(f"Erro: Não foi possível obter access_token para '{nome_atleta}'.")
        continue

    headers = {'Authorization': f'Bearer {access_token}'}
    url = "https://www.strava.com/api/v3/athlete/activities"
    
    resposta = requests.get(url, headers=headers, params={'per_page': 200, 'page': 1})
    
    if resposta.status_code == 200:
        atividades = resposta.json()
        km_total = 0.0
        alt_total = 0.0
        treinos = 0
        
        for act in atividades:
            tipo = act.get('type')
            data_inicio = act.get('start_date', '')
            
            if tipo in ['Run', 'TrailRun'] and data_inicio.startswith('2026'):
                dist_km = act.get('distance', 0.0) / 1000.0
                alt = act.get('total_elevation_gain', 0.0)
                
                km_total += dist_km
                alt_total += alt
                treinos += 1

        dados_ranking.append({
            'Atleta': nome_atleta,
            'KM Total Bruto': km_total,
            'KM Total': formatar_km(km_total),
            'Altimetria (m)': formatar_alt(alt_total),
            'Treinos': treinos
        })
        print(f"✓ {nome_atleta}: {formatar_km(km_total)} ({treinos} treinos)")
    else:
        print(f"Erro na consulta do Strava para {nome_atleta}: Status {resposta.status_code}")

if dados_ranking:
    df = pd.DataFrame(dados_ranking)
    df = df.sort_values(by='KM Total Bruto', ascending=False)
    df = df.drop(columns=['KM Total Bruto'])
else:
    df = pd.DataFrame(columns=['Atleta', 'KM Total', 'Altimetria (m)', 'Treinos'])

with pd.ExcelWriter(NOME_ARQUIVO, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Ranking CNB 2026', index=False)

print(f"Sincronização da planilha {NOME_ARQUIVO} concluída com sucesso!")
