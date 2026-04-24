import os
import pandas as pd
import json

# Configuração
SOURCE_DIR = r"C:\Users\Cassyano\OneDrive - Brasil do Parafusos\Comprasbrasil - Compras\Entradas Mensal"
OUTPUT_FILE = "data-entradas.js"

def sync_data():
    all_data = []
    
    if not os.path.exists(SOURCE_DIR):
        print(f"Erro: Pasta não encontrada: {SOURCE_DIR}")
        return

    files = [f for f in os.listdir(SOURCE_DIR) if f.endswith(('.xlsx', '.xls'))]
    
    if not files:
        print("Nenhum arquivo Excel encontrado na pasta.")
        return

    for file in files:
        path = os.path.join(SOURCE_DIR, file)
        try:
            # Ler Excel
            df = pd.read_excel(path)
            
            # Normalizar colunas
            # Buscamos: Dt.movto, Razão social, Vlr.cont
            col_map = {
                'Dt.movto': None,
                'Razão social': None,
                'Vlr.cont': None
            }
            
            for col in df.columns:
                c = str(col).strip().lower()
                if c == 'dt.movto': col_map['Dt.movto'] = col
                if c == 'razão social': col_map['Razão social'] = col
                if c == 'vlr.cont': col_map['Vlr.cont'] = col
            
            # Se não achou com nome exato, tenta parcial
            for col in df.columns:
                c = str(col).strip().lower()
                if not col_map['Dt.movto'] and 'movto' in c: col_map['Dt.movto'] = col
                if not col_map['Razão social'] and ('razão' in c or 'fornecedor' in c): col_map['Razão social'] = col
                if not col_map['Vlr.cont'] and ('vlr' in c or 'valor' in c): col_map['Vlr.cont'] = col

            if all(col_map.values()):
                # Extrair dados
                for _, row in df.iterrows():
                    val = row[col_map['Vlr.cont']]
                    if pd.isna(val) or val <= 0: continue
                    
                    dt = row[col_map['Dt.movto']]
                    if pd.isna(dt): continue
                    
                    # Converter data para string ISO ou objeto
                    if isinstance(dt, pd.Timestamp):
                        dt_str = dt.strftime('%Y-%m-%d')
                    else:
                        dt_str = str(dt)

                    all_data.append({
                        'date': dt_str,
                        'supplier': str(row[col_map['Razão social']]).strip(),
                        'value': float(val)
                    })
                print(f"Processado: {file} ({len(df)} linhas)")
            else:
                print(f"Aviso: Colunas não identificadas em {file}. Esperado: Dt.movto, Razão social, Vlr.cont")
                
        except Exception as e:
            print(f"Erro ao processar {file}: {e}")

    # Salvar como JS para evitar problemas de CORS no navegador
    js_content = f"const PRE_LOADED_ENTRADAS = {json.dumps(all_data, indent=4)};"
    with open(OUTPUT_FILE, 'w', encoding='utf-8') as f:
        f.write(js_content)
    
    print(f"Sincronização concluída! {len(all_data)} registros salvos em {OUTPUT_FILE}")

if __name__ == "__main__":
    sync_data()
