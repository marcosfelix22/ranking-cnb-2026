import os
import requests
import pandas as pd

CLIENT_ID = (os.environ.get('CLIENT_ID') or '').strip()
CLIENT_SECRET = (os.environ.get('CLIENT_SECRET') or '').strip()
TOKEN_MARCOS = (os.environ.get('TOKEN_MARCOS') or '').strip()

# Cole o novo 'code' da Juliana copiado da URL aqui:
CODIGO_JULIANA = "8a2b8681422d2e8b686c23c1cec8cc002fdb18f7"

print("--- EXECUTANDO TROCA DIRETA E ATUALIZAÇÃO DO RANKING ---")

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

def obter_dados_atleta(access_token):
    headers = {'Authorization': f'Bearer {access_token}'}
    res = requests.get("https://www.strava.com/api/v3/athlete/activities", headers=headers, params={'per_page': 200})
    if res.status_code == 200:
        km_total, alt_total, treinos = 0.0, 0.0, 0
        for act in res.json():
            if act.get('type') in ['Run', 'TrailRun'] and act.get('start_date', '').startswith('2026'):
                km_total += act.get('distance', 0.0) / 1000.0
                alt_total += act.get('total_elevation_gain', 0.0)
                treinos += 1
        return km_total, alt_total, treinos
    return None, None, None

dados_ranking = []

# 1. Processa Marcos Felix via Refresh Token
token_marcos = obter_access_token(TOKEN_MARCOS)
if token_marcos:
    km, alt, treinos = obter_dados_atleta(token_marcos)
    if km is not None:
        dados_ranking.append({'Atleta': 'Marcos Felix', 'KM Total Bruto': km, 'KM Total': f"{km:,.1f}".replace(".", ",") + " km", 'Altimetria (m)': f"{int(alt):,} m".replace(",", "."), 'Treinos': treinos})
        print(f"✓ Marcos Felix: {km:.1f} km em {treinos} treinos")

# 2. Processa Juliana Nogueira trocando o 'code' em tempo de execução
payload_troca = {
    'client_id': CLIENT_ID,
    'client_secret': CLIENT_SECRET,
    'code': CODIGO_JULIANA,
    'grant_type': 'authorization_code'
}
res_troca = requests.post("https://www.strava.com/oauth/token", data=payload_troca)

if res_troca.status_code == 200:
    dados_juliana = res_troca.json()
    access_token_juliana = dados_juliana.get('access_token')
    refresh_token_juliana = dados_juliana.get('refresh_token')
    
    print("✅ Troca do código da Juliana realizada com sucesso!")
    print(f"🔑 REFRESH TOKEN OFICIAL DA JULIANA: {refresh_token_juliana}")
    
    km_j, alt_j, treinos_j = obter_dados_atleta(access_token_juliana)
    if km_j is not None:
        dados_ranking.append({'Atleta': 'Juliana Nogueira', 'KM Total Bruto': km_j, 'KM Total': f"{km_j:,.1f}".replace(".", ",") + " km", 'Altimetria (m)': f"{int(alt_j):,} m".replace(",", "."), 'Treinos': treinos_j})
        print(f"✓ Juliana Nogueira: {km_j:.1f} km em {treinos_j} treinos")
else:
    print(f"❌ Erro na troca do código da Juliana: Status {res_troca.status_code} - {res_troca.text}")

# 3. Monta e salva o Ranking
if dados_ranking:
    df = pd.DataFrame(dados_ranking)
    df = df.sort_values(by='KM Total Bruto', ascending=False).drop(columns=['KM Total Bruto'])
    with pd.ExcelWriter('Ranking_CNB_2026.xlsx', engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Ranking CNB 2026', index=False)
    print("✅ Planilha Ranking_CNB_2026.xlsx gerada e atualizada com sucesso!")
