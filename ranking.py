import os
import requests
import pandas as pd

# --- CONFIGURAÇÕES DA API DO STRAVA ---
CLIENT_ID = (os.environ.get('CLIENT_ID') or '').strip()
CLIENT_SECRET = (os.environ.get('CLIENT_SECRET') or '').strip()
NOME_ARQUIVO = 'Ranking_CNB_2026.xlsx'

# Dicionário com os atletas e seus respectivos tokens salvos nos GitHub Secrets
ATLETAS = {
    "Marcos Felix": os.environ.get('TOKEN_MARCOS'),
    "Juliana Nogueira": os.environ.get('TOKEN_JULIANA'),
    "Cristiano Silva": os.environ.get('TOKEN_CRISTIANO_SILVA'),
    "Fabio Oliveira": os.environ.get('TOKEN_FABIO'),
}

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
        if res.status_code == 200:
            return res.json().get('access_token')
        else:
            print(f"❌ Erro ao renovar token no Strava: {res.status_code}")
            return None
    except Exception as e:
        print(f"❌ Exceção ao obter access_token: {e}")
        return None

def formatar_km(valor):
    return f"{valor:,.1f}".replace(",", "X").replace(".", ",").replace("X", ".") + " km"

def formatar_alt(valor):
    return f"{int(valor):,}".replace(",", ".") + " m"

print("--- EXECUTANDO ATUALIZAÇÃO AUTOMÁTICA DO RANKING CNB 2026 ---")

dados_ranking = []

for nome_atleta, ref_token in ATLETAS.items():
    if not ref_token:
        print(f"⚠️ Secret/Token do atleta '{nome_atleta}' não encontrado. Pulando...")
        continue

    access_token = obter_access_token(ref_token)
    if not access_token:
        print(f"❌ Não foi possível autorizar o acesso para '{nome_atleta}'.")
        continue

    headers = {'Authorization': f'Bearer {access_token}'}
    url = "https://www.strava.com/api/v3/athlete/activities"
    
    # Consulta até 200 atividades recentes do atleta
    res = requests.get(url, headers=headers, params={'per_page': 200, 'page': 1})
    
    if res.status_code == 200:
        atividades = res.json()
        km_total = 0.0
        alt_total = 0.0
        treinos = 0
        
        for act in atividades:
            tipo = act.get('type')
            # Usando a data local (fuso horário do atleta)
            data_inicio = act.get('start_date_local', '')
            
            # Filtra apenas corridas e corridas de trilha em 2026
            if tipo in ['Run', 'TrailRun'] and data_inicio.startswith('2026'):
                dist_km = act.get('distance', 0.0) / 1000.0
                alt = act.get('total_elevation_gain', 0.0)
                
                # Debug para Marcos Felix: mostra todos os treinos no log do GitHub Actions
                if nome_atleta == "Marcos Felix":
                    print(f"  └─ [{data_inicio[:10]}] {act.get('name')}: {dist_km:.2f} km")
                
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
        print(f"✓ {nome_atleta}: {formatar_km(km_total)} em {treinos} treinos")
    else:
        print(f"❌ Erro na consulta do Strava para {nome_atleta}: Status {res.status_code}")

# Monta o Ranking ordenado pelo KM Total (do maior para o menor)
if dados_ranking:
    df = pd.DataFrame(dados_ranking)
    df = df.sort_values(by='KM Total Bruto', ascending=False)
    df = df.drop(columns=['KM Total Bruto'])
else:
    df = pd.DataFrame(columns=['Atleta', 'KM Total', 'Altimetria (m)', 'Treinos'])

# Salva na planilha Excel
with pd.ExcelWriter(NOME_ARQUIVO, engine='openpyxl') as writer:
    df.to_excel(writer, sheet_name='Ranking CNB 2026', index=False)

print(f"✅ Sincronização da planilha {NOME_ARQUIVO} concluída com sucesso!")
