import os
import requests
import pandas as pd

CLIENT_ID = (os.environ.get('CLIENT_ID') or '').strip()
CLIENT_SECRET = (os.environ.get('CLIENT_SECRET') or '').strip()

# Cole o seu 'code' novinho recém-copiado do navegador aqui:
CODIGO_AUTORIZACAO = "bf76f0244188a1a85989dea5323e592b7ce19c58" 

print("--- EXECUTANDO AUTENTICAÇÃO E CONSOLIDAÇÃO DO RANKING ---")

payload_troca = {
    'client_id': CLIENT_ID,
    'client_secret': CLIENT_SECRET,
    'code': CODIGO_AUTORIZACAO,
    'grant_type': 'authorization_code'
}

res_troca = requests.post("https://www.strava.com/oauth/token", data=payload_troca)

if res_troca.status_code == 200:
    dados_oauth = res_troca.json()
    access_token = dados_oauth.get('access_token')
    refresh_token_oficial = dados_oauth.get('refresh_token')
    
    print("✅ Autenticação realizada com sucesso!")
    print(f"🔑 SEU REFRESH TOKEN OFICIAL COM ESCOPO TOTAL É: {refresh_token_oficial}")
    
    headers = {'Authorization': f'Bearer {access_token}'}
    res_act = requests.get("https://www.strava.com/api/v3/athlete/activities", headers=headers, params={'per_page': 200})
    
    if res_act.status_code == 200:
        atividades = res_act.json()
        km_total = 0.0
        alt_total = 0.0
        treinos = 0
        
        for act in atividades:
            tipo = act.get('type')
            data_inicio = act.get('start_date', '')
            if tipo in ['Run', 'TrailRun'] and data_inicio.startswith('2026'):
                km_total += act.get('distance', 0.0) / 1000.0
                alt_total += act.get('total_elevation_gain', 0.0)
                treinos += 1
                
        print(f"✓ Marcos Felix: {km_total:.1f} km em {treinos} treinos.")
        
        df = pd.DataFrame([{
            'Atleta': 'Marcos Felix', 
            'KM Total': f"{km_total:,.1f}".replace(".", ",") + " km", 
            'Altimetria (m)': f"{int(alt_total):,} m".replace(",", "."), 
            'Treinos': treinos
        }])
        
        with pd.ExcelWriter('Ranking_CNB_2026.xlsx', engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Ranking CNB 2026', index=False)
            
        print("✅ Planilha Ranking_CNB_2026.xlsx gerada e atualizada com sucesso!")
    else:
        print(f"❌ Erro na consulta das atividades: Status {res_act.status_code} - {res_act.text}")
else:
    print(f"❌ Erro na troca do código: Status {res_troca.status_code} - {res_troca.text}")
