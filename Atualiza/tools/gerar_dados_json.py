"""
Utilitário para regenerar data/dados.json a partir da planilha Musica.xlsx
(aba "Ensaios"), fora da interface web. Uso:

    pip install pandas openpyxl
    python gerar_dados_json.py caminho/para/Musica.xlsx

O arquivo gerado é escrito em ../data/dados.json.

Se a planilha tiver uma coluna "Cidade", ela é usada diretamente. Se não
tiver (planilhas antigas), o script tenta inferir a cidade a partir do
Setor/Congregação como fallback.
"""
import sys, json, re, unicodedata
import pandas as pd

MESES = ["Janeiro","Fevereiro","Março","Abril","Maio","Junho","Julho",
         "Agosto","Setembro","Outubro","Novembro","Dezembro"]
MES_MAP = {m.upper(): i+1 for i, m in enumerate(MESES)}
DIAS = {"domingo":0,"segunda-feira":1,"terca-feira":2,"terça-feira":2,
        "quarta-feira":3,"quinta-feira":4,"sexta-feira":5,"sabado":6,"sábado":6}
DIA_LABEL = {0:"Domingo",1:"Segunda-feira",2:"Terça-feira",3:"Quarta-feira",
             4:"Quinta-feira",5:"Sexta-feira",6:"Sábado"}
# normaliza variações de grafia encontradas na planilha para um rótulo padrão.
# tipos que não estiverem aqui são mantidos como estão (title case), então
# qualquer tipo novo que a igreja venha a criar no futuro continua funcionando.
TIPO_MAP = {
    "MENSAL":"Mensal", "BIMESTRAL":"Bimestral",
    "EXTRAS":"Extra", "EXTRA":"Extra",
    "QUATRIMESTRAIS":"Quadrimestral", "QUATRIMESTRAL":"Quadrimestral", "QUADRIMESTRAL":"Quadrimestral",
    "SEMENTRAL":"Semestral", "SEMESTRAL":"Semestral",
    "TRIMESTRAIS":"Trimestral", "TRIMESTRAL":"Trimestral",
    "ENSAIO REGIONAL":"Ensaio Regional",
    "TESTE E EXAMES":"Teste e Exames", "TESTES E EXAMES":"Teste e Exames",
}
SETORES_TERESINA = {"Setor 1","Setor 2","Setor 3","Setor 4","Setor 5","Setor 10","ADM União"}

def strip_accents(s):
    return ''.join(c for c in unicodedata.normalize('NFD', s) if unicodedata.category(c) != 'Mn')

def parse_dia_semana(s):
    m = re.match(r'^(\d)ª?\s*(.+)$', s.strip())
    ordinal = int(m.group(1))
    raw = m.group(2).strip().lower()
    idx = DIAS.get(raw, DIAS.get(strip_accents(raw)))
    return ordinal, idx

def normalizar_tipo(raw):
    raw = raw.strip()
    mapped = TIPO_MAP.get(raw.upper())
    if mapped:
        return mapped
    return " ".join(w.capitalize() for w in raw.split())

def cidade_fallback(setor, congregacao):
    """Usado apenas se a planilha não tiver coluna Cidade."""
    if setor in SETORES_TERESINA: return "Teresina"
    if setor == "Setor Floriano": return "Floriano" if congregacao == "Floriano Central" else congregacao
    if setor == "Setor Parnaiba": return "Parnaíba" if congregacao == "Parnaiba Central" else congregacao
    if setor == "Setor Campo Maior": return congregacao
    if setor == "Setor Timon": return "Timon"
    return congregacao

def main(path):
    df = pd.read_excel(path, sheet_name='Ensaios')
    tem_coluna_cidade = "Cidade" in df.columns
    records = []
    for i, row in df.iterrows():
        ordinal, dia_idx = parse_dia_semana(row['Dia_Semana'])
        mes_num = MES_MAP[row['mes'].strip().upper()]
        tipo = normalizar_tipo(str(row['TIPO']))
        h = row['HORARIO']
        obs = row['OBS'] if pd.notna(row['OBS']) else ""
        congregacao = str(row['Congregação']).strip()
        setor = str(row['Setor']).strip()
        if tem_coluna_cidade and pd.notna(row['Cidade']) and str(row['Cidade']).strip():
            cidade = str(row['Cidade']).strip()
        else:
            cidade = cidade_fallback(setor, congregacao)
        records.append({
            "id": i+1, "mesNumero": mes_num, "mes": MESES[mes_num-1],
            "diaSemanaOrdinal": ordinal, "diaSemanaIndice": dia_idx,
            "diaSemanaLabel": DIA_LABEL[dia_idx], "diaSemanaTexto": f"{ordinal}ª {DIA_LABEL[dia_idx]}",
            "congregacao": congregacao,
            "cidade": cidade,
            "setor": setor, "tipo": tipo,
            "horario": f"{h.hour:02d}:{h.minute:02d}",
            "encarregadoLocal": str(row['Encarregado Local']).strip(),
            "encarregadoRegional": str(row['Encarregado Regional']).strip(),
            "observacoes": obs.strip() if isinstance(obs, str) else ""
        })
    tipos_presentes = sorted(set(r["tipo"] for r in records))
    out = {"meta": {"totalRegistros": len(records), "fonte": path,
                     "tiposEnsaio": tipos_presentes},
           "ensaios": records}
    with open('../data/dados.json', 'w', encoding='utf-8') as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"OK: {len(records)} ensaios gravados em ../data/dados.json")
    print(f"Tipos encontrados: {tipos_presentes}")
    print(f"Coluna Cidade na planilha: {'sim' if tem_coluna_cidade else 'não (usado fallback)'}")

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Uso: python gerar_dados_json.py caminho/para/Musica.xlsx"); sys.exit(1)
    main(sys.argv[1])
