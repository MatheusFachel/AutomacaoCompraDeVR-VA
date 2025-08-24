# No arquivo main.py
# SUBSTITUA APENAS A FUNÇÃO carregar_planilhas PELA VERSÃO ABAIXO

def carregar_planilhas(caminho_pasta_entrada: Path) -> Optional[Dict[str, pd.DataFrame]]:
    if not caminho_pasta_entrada.is_dir():
        raise FileNotFoundError(f"O diretório de entrada não foi encontrado em '{caminho_pasta_entrada}'")

    mapa_arquivos = {
        "ATIVOS.xlsx": "ativos", "Base dias uteis.xlsx": "dias_uteis",
        "Base sindicato x valor.xlsx": "sindicato_valor", "FÉRIAS.xlsx": "ferias",
        "ADMISSÃO ABRIL.xlsx": "admissao", "DESLIGADOS.xlsx": "desligados",
        "AFASTAMENTOS.xlsx": "afastamentos", "APRENDIZ.xlsx": "aprendiz",
        "EXTERIOR.xlsx": "exterior", "ESTÁGIO.xlsx": "estagio",
        "VR MENSAL 05.2025.xlsx": "template_final_e_validacoes"
    }
    dataframes = {}
    print("--- Módulo 1: Iniciando carregamento e limpeza das planilhas ---")
    for nome_arquivo, chave in mapa_arquivos.items():
        caminho_arquivo = caminho_pasta_entrada / nome_arquivo
        try:
            if nome_arquivo == "VR MENSAL 05.2025.xlsx":
                abas = pd.read_excel(caminho_arquivo, sheet_name=["VR MENSAL 05.2025", "Validações"], header=0)
                dataframes["template_final"] = abas["VR MENSAL 05.2025"]
                dataframes["validacoes"] = abas["Validações"]
                print(f"  - Arquivo '{nome_arquivo}' (2 abas) carregado.")
            else:
                # CORREÇÃO APLICADA AQUI -> header=0
                df_temp = pd.read_excel(caminho_arquivo, header=0)
                dataframes[chave] = df_temp
                print(f"  - Arquivo '{nome_arquivo}' carregado.")
        except FileNotFoundError:
            print(f"  - AVISO: O arquivo '{nome_arquivo}' não foi encontrado.")
        except Exception as e:
            raise IOError(f"Falha ao ler o arquivo '{nome_arquivo}'. Erro: {e}")

    for chave, df in dataframes.items():
        colunas_limpas = [re.sub(r'\s+', ' ', str(col)).strip().upper() for col in df.columns]
        df.columns = colunas_limpas
    print("  - Todas as colunas foram padronizadas.")
    return dataframes