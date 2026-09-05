import urllib.request

# Função com cabeçalho de navegador e fallback automático
def carregar_google_sheets(url):
    id_match = re.search(r"/d/([a-zA-Z0-9-_]+)", url)
    if not id_match:
        return None
    sheet_id = id_match.group(1)
    
    gid_match = re.search(r"[#&?]gid=([0-9]+)", url)
    gid_param = f"&gid={gid_match.group(1)}" if gid_match else ""
    
    timestamp = int(time.time() * 1000)
    
    # Lista com as duas rotas oficiais do Google Sheets
    urls_tentativas = [
        f"https://docs.google.com/spreadsheets/d/{sheet_id}/gviz/tq?tqx=out:csv{gid_param}&t={timestamp}",
        f"https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=csv{gid_param}&t={timestamp}"
    ]
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    
    for link in urls_tentativas:
        try:
            req = urllib.request.Request(link, headers=headers)
            with urllib.request.urlopen(req, timeout=10) as resp:
                conteudo = resp.read()
                df_lido = pd.read_csv(io.BytesIO(conteudo))
                if df_lido is not None and not df_lido.empty:
                    return df_lido
        except Exception:
            continue
            
    return None

