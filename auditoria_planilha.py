#!/usr/bin/env python3
"""
AUDITORIA FORENSE — Planilha Orçamento V Preço de Venda (.xlsx)
Infinity Health, SA · Concursos IIAM 35A003641 · Rev.0

Uso:
    python3 auditoria_planilha.py "Planilha Orçamento V Preço de Venda.xlsx" [--out achados.csv]
    python3 auditoria_planilha.py --self-test        # valida o motor com workbook sintético (erros plantados)

O motor:
  FASE 1 — estrutura: abas, dimensões, cabeçalhos detectados, formatações condicionais vermelhas.
  FASE 2 — mapeamento de colunas por palavras-chave (PT).
  FASE 3 — controlos T1..T14 / C1..C10 com endereço exacto ABA!CÉLULA e impacto em MT.
  FASE 4 — relatório consola + CSV.

Notas metodológicas (Moçambique):
  - Bandas MFN: 0 / 2.5 / 5 / 7.5 / 20% sobre CIF; IVA-import 16% sobre (CIF+direitos+ICE).
  - Seguro-ref: 2% x (FOB_MZN + Frete_MZN). Taxa bancária-ref: 5.100 MZN 1x/processo.
  - IVA recuperável (regime normal) NÃO entra na base de markup.
"""
import sys, csv, re
from collections import Counter, defaultdict
from openpyxl import load_workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule

# ---------------------------------------------------------------- referência HS (Moçambique, indicativa)
HS_REF = {
    # hs: (taxa_mfn, categoria, subcategoria, confiança)
    "3822": (0.0, "Reagente", "diagnóstico/lab (frio ou não-frio: taxa igual, muda o frete)", "alta"),
    "3002": (0.0, "Reagente", "antisoros/frações de sangue", "média"),
    "3003": (0.0, "Reagente", "medicamentos doseamento", "média"),
    "3004": (0.0, "Reagente", "medicamentos acondicionados", "média"),
    "3926": (20.0, "Consumível", "plásticos (confirmar 7.5 vs 20 c/ despachante)", "confirmar"),
    "7017": (5.0, "Consumível", "vidro lab", "média"),
    "4016": (7.5, "Consumível", "borracha (confirmar 7.5 vs 20)", "confirmar"),
    "7326": (20.0, "Consumível", "metal acabados (confirmar 7.5 vs 20)", "confirmar"),
    "7323": (20.0, "Consumível", "utensílios domésticos aço (risco p/ 'panelas')", "alta"),
    "8414": (5.0, "Equipamento", "capelas/hoods, cabines biossegurança", "média"),
    "8415": (20.0, "Equipamento", "ar condicionado (consumo)", "alta"),
    "8418": (20.0, "Equipamento", "refrigeração/congelação (consumo)", "alta"),
    "8419": (5.0, "Equipamento", "esterilizadores/secadores/incubadoras (8419.20 pode ser 0% médico)", "média"),
    "8421": (5.0, "Equipamento", "centrífugas", "alta"),
    "8422": (5.0, "Equipamento", "encher/fechar/rotular/lavar", "alta"),
    "8479": (5.0, "Equipamento", "agitadores/mistura + partes (8479.90)", "alta"),
    "9011": (5.0, "Equipamento", "microscópios (verificar 2.5 instr. científicos)", "média"),
    "9025": (5.0, "Equipamento", "termómetros/higrómetros", "média"),
    "SERV": (0.0, "Serviço", "treino/instalação/calibração: 0% direitos, IVA 16% na factura", "alta"),
}
BANDAS = {0.0, 2.5, 5.0, 7.5, 20.0}
TOL = 0.02  # tolerância MT p/ comparações

# ---------------------------------------------------------------- utilidades
def norm(s):
    if s is None: return ""
    s = str(s).strip().lower().replace("_", " ").replace("-", " ")
    for a, b in [("ã","a"),("õ","o"),("á","a"),("à","a"),("â","a"),("é","e"),("ê","e"),
                 ("í","i"),("ó","o"),("ô","o"),("ú","u"),("ç","c")]:
        s = s.replace(a, b)
    return re.sub(r"\s+", " ", s)

HEADER_KEYS = {  # chave_lógica: [palavras-chave no cabeçalho]
    "item": ["item", "#", "nº", "num"],
    "desc": ["descricao", "descrição", "designacao", "produto", "bem"],
    "qtd": ["qtd", "quant", "quantidade"],
    "moeda": ["moeda", "currency", "divisa"],
    "fob": ["fob", "preco fob", "custo fob", "valor fob", "preco fornecedor", "cotacao"],
    "cambio": ["cambio", "câmbio", "tx cambio", "taxa cambio", "rate"],
    "fob_mzn": ["fob mzn", "fob (mzn)", "custo mzn", "valor mzn", "fob mt"],
    "frete": ["frete", "freight", "transporte internacional"],
    "seguro": ["seguro", "insurance"],
    "cif": ["cif"],
    "hs": ["hs", "sh code", "pauta", "posicao pautal", "ncm"],
    "taxa": ["taxa", "direito", "duty", "tarifa", "% imposto", "di "],
    "direitos": ["direitos", "ii ", "imposto import"],
    "iva": ["iva", "vat"],
    "despacho": ["despacho", "desembaraco", "desalfandeg", "clearing"],
    "bancaria": ["bancaria", "bancária", "taxa banca", "comissao banca"],
    "markup": ["markup", "margem", "margin", "majoracao"],
    "pu": ["preco unit", "p.u.", "pv unit", "venda unit"],
    "total": ["preco total", "total s/iva", "total sem iva", "subtotal"],
    "categoria": ["categoria", "classe", "tipo produto", "grupo"],
    "origem": ["origem", "pais", "country"],
    "modofrete": ["via", "modal", "tipo frete", "aereo", "maritimo", "rodoviario"],
}
MOEDAS = ["USD", "GBP", "ZAR", "CNY", "EUR", "MZN", "USD$", "R ", "€", "£"]

class Achado:
    def __init__(self, seq, sev, local, teste, achado, causa, impacto_mt, correcao, ambito):
        self.seq, self.sev, self.local, self.teste = seq, sev, local, teste
        self.achado, self.causa = achado, causa
        self.impacto_mt, self.correcao, self.ambito = impacto_mt, correcao, ambito
    def row(self):
        return [self.seq, self.sev, self.local, self.teste, self.achado, self.causa,
                f"{self.impacto_mt:.2f}" if isinstance(self.impacto_mt, float) else self.impacto_mt,
                self.correcao, self.ambito]

def is_red_fill(cell):
    try:
        f = cell.fill
        if f and f.fgColor and f.fgColor.rgb and f.fgColor.rgb != "00000000":
            rgb = str(f.fgColor.rgb).upper()
            return rgb in ("FFFF0000", "00FF0000", "FFFFC7CE", "00FFC7CE") or rgb.endswith("FF0000") or rgb.endswith("C7CE")
    except Exception:
        pass
    return False

def num(v):
    if v is None or isinstance(v, bool): return None
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip().replace(" ", "")
    if not s: return None
    # 1.234.567,89 (PT) vs 1234567.89
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try: return float(s)
    except ValueError: return None

def pct(v):
    n = num(v)
    if n is None: return None
    return n * 100.0 if abs(n) < 1.0 and n != 0 else n  # 0.16 -> 16

# ---------------------------------------------------------------- FASE 1+2: estrutura e mapeamento
def detectar_cabecalho(ws):
    best, best_hit = None, 0
    for r in range(1, min(25, ws.max_row + 1)):
        vals = [norm(ws.cell(r, c).value) for c in range(1, min(60, ws.max_column + 1))]
        hits = sum(1 for v in vals for ks in HEADER_KEYS.values() for k in ks if k and k in v)
        if hits > best_hit:
            best_hit, best = hits, r
    return best, best_hit

def mapear_colunas(ws, hr):
    mapa = {}
    for c in range(1, ws.max_column + 1):
        v = norm(ws.cell(hr, c).value)
        if not v: continue
        for chave, kws in HEADER_KEYS.items():
            if chave in mapa: continue
            if any(k in v for k in kws):
                mapa[chave] = c
    return mapa

def condicoes_vermelhas(wb):
    out = []
    for ws in wb.worksheets:
        try: cfs = ws.conditional_formatting._cf_rules
        except Exception: continue
        for rng, rules in cfs.items():
            for rl in rules:
                txt = ""
                try: txt = str(rl.dxf.fill.fgColor.rgb) if rl.dxf and rl.dxf.fill else ""
                except Exception: pass
                f = str(getattr(rl, "formula", "") or "") + str(getattr(rl, "text", "") or "")
                if "FF0000" in txt.upper() or "C7CE" in txt.upper():
                    out.append((ws.title, str(rng), f[:120]))
    return out

def taxas_declaradas(wb):
    """Procura bloco de câmbios declarados (ex.: 'USD 63.90'). Retorna {MOEDA: (taxa, local)}."""
    out = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows(min_row=1, max_row=40, max_col=20):
            for i, cell in enumerate(row):
                t = norm(cell.value)
                for m in ["USD", "GBP", "ZAR", "CNY", "EUR"]:
                    if t == m.lower() or t.startswith(m.lower() + " ") or t.startswith(m.lower() + "/"):
                        for nb in row[i+1:i+4]:
                            n = num(nb.value)
                            if n and 0.01 < n < 100000:
                                out.setdefault(m, (n, f"{ws.title}!{nb.coordinate}"))
                                break
    return out

# ---------------------------------------------------------------- FASE 3: controlos
def auditar(path, out_csv=None):
    wb = load_workbook(path, data_only=False)
    wbv = load_workbook(path, data_only=True)  # valores calculados
    ach, seq = [], [0]
    def add(sev, local, teste, achado, causa, impacto, correcao, ambito):
        seq[0] += 1
        ach.append(Achado(seq[0], sev, local, teste, achado, causa, impacto, correcao, ambito))

    print(f"\n{'='*78}\nAUDITORIA: {path}\nAbas: {[s.title for s in wb.worksheets]}")
    decl = taxas_declaradas(wb)
    print(f"Câmbios declarados detectados: {decl if decl else 'NENHUM (T2 incompleto — ver A-T2)'}")
    if not decl:
        add("ALTA", "workbook!A1:Z40", "T2-fonte-cambio",
            "Bloco de câmbios declarados não detectado nas 40 primeiras linhas.",
            "Taxas hardcoded por linha ou bloco fora da área varrida.",
            "n/d", "Criar aba PRESSUPOSTOS com USD/GBP/ZAR/CNY/EUR + fonte(BM/banco)+data+spread; apontar fórmulas para lá.", "planilha")

    for ws in wb.worksheets:
        wsv = wbv[ws.title]
        hr, hits = detectar_cabecalho(ws)
        if not hr or hits < 3:
            print(f"[{ws.title}] sem tabela de itens detectada (hits={hits}) — salto controlos de linha.")
            continue
        mapa = mapear_colunas(ws, hr)
        print(f"[{ws.title}] cabeçalho linha {hr} → colunas: {mapa}")
        if len(mapa) < 4:
            add("MEDIA", f"{ws.title}!{hr}:{hr}", "MAPEAMENTO",
                f"Só {len(mapa)} colunas reconhecidas — cabeçalhos fora do padrão.",
                "Nomenclatura de cabeçalho inesperada.", "n/d",
                "Normalizar cabeçalhos (Item/Descrição/QTD/Moeda/FOB/Câmbio/Frete/Seguro/CIF/HS/Taxa/IVA/Despacho/Markup/PU/Total).", "aba")
            continue
        C = lambda k, r: wsv.cell(r, mapa[k]).value if k in mapa else None
        Cf = lambda k, r: ws.cell(r, mapa[k]).coordinate if k in mapa else "?"
        L = lambda k, r: f"{ws.title}!{Cf(k, r)}"
        nrows = ws.max_row

        # --- T1/C1: câmbio único por moeda + conformidade com declarado
        if "moeda" in mapa and "cambio" in mapa:
            por_moeda = defaultdict(list)
            for r in range(hr + 1, nrows + 1):
                m = str(C("moeda", r) or "").upper().strip()
                m = {"DOLAR": "USD", "DÓLAR": "USD", "LIBRA": "GBP", "RAND": "ZAR",
                     "YUAN": "CNY", "EURO": "EUR", "METICAL": "MZN"}.get(m, m)
                tx = num(C("cambio", r))
                if m and tx: por_moeda[m].append((tx, r))
            for m, lst in por_moeda.items():
                taxas = {round(t, 6) for t, _ in lst}
                if len(taxas) > 1:
                    detalhe = ", ".join(f"L{r}={t}" for t, r in sorted(lst)[:8])
                    add("ALTA", f"{ws.title}!col-câmbio({mapa['cambio']})", "T1-cambio-unico",
                        f"Moeda {m} com {len(taxas)} taxas distintas: {sorted(taxas)} ({detalhe}…).",
                        "Taxa digitada à mão em algumas linhas.", "modelo",
                        f"Unificar {m} numa célula-nome (ex.: CAMBIO_{m}); ={sorted(taxas)[0]} provisório até confirmar fonte.", "planilha")
                if m in decl and lst:
                    med = sum(t for t, _ in lst) / len(lst)
                    if abs(med - decl[m][0]) / decl[m][0] > 0.0005:
                        add("ALTA", L("cambio", lst[0][1]), "T2-conversao-vs-declarado",
                            f"{m}: taxa usada {med:.4f} ≠ declarada {decl[m][0]:.4f} ({decl[m][1]}).",
                            "Linhas não apontam ao bloco declarado.", round(abs(med-decl[m][0]), 4),
                            f"Apontar câmbios {m} para {decl[m][1]}.", "planilha")

        # --- C2: FOB x câmbio = FOB_MZN
        if {"fob", "cambio", "fob_mzn"} <= set(mapa):
            for r in range(hr + 1, nrows + 1):
                if C("desc", r) is None and C("item", r) is None: continue
                fob, tx, reg = num(C("fob", r)), num(C("cambio", r)), num(C("fob_mzn", r))
                if fob is not None and tx is not None and reg is not None:
                    esp = round(fob * tx, 2)
                    if abs(esp - reg) > max(TOL, abs(reg) * 0.001):
                        qtd = num(C("qtd", r)) or 1
                        add("ALTA", L("fob_mzn", r), "C2-conversao",
                            f"L{r}: FOB×câmbio = {esp:,.2f} ≠ registado {reg:,.2f} (dif. {esp-reg:,.2f}).",
                            "Câmbio errado, FOB errado ou fórmula quebrada.", round((esp-reg)*1.0, 2),
                            f"Corrigir para {esp:,.2f} (= {fob} × {tx}).", "item")

        # --- C3: seguro 2%
        if "seguro" in mapa:
            for r in range(hr + 1, nrows + 1):
                if C("desc", r) is None and C("item", r) is None: continue
                seg = num(C("seguro", r))
                if seg is None: continue
                fobm = num(C("fob_mzn", r)) if "fob_mzn" in mapa else None
                fre = num(C("frete", r)) if "frete" in mapa else 0.0
                if fobm is not None:
                    base = fobm + (fre or 0.0)
                    esp = round(base * 0.02, 2)
                    if base > 0 and abs(esp - seg) > max(TOL, esp * 0.02):
                        add("MEDIA", L("seguro", r), "T13-seguro-2pct",
                            f"L{r}: seguro {seg:,.2f} ≠ 2%×(FOB+Frete)={esp:,.2f} (base {base:,.2f}).",
                            "Base errada (só FOB / sobre CIF / % divergente).", round(esp-seg, 2),
                            f"=ROUND(({Cf('fob_mzn', r)}+{Cf('frete', r) if 'frete' in mapa else 0})*2%;2) → {esp:,.2f}.", "item")

        # --- T12: taxa bancária 5.100
        if "bancaria" in mapa:
            hits5100 = [(r, num(C("bancaria", r))) for r in range(hr + 1, nrows + 1)
                        if num(C("bancaria", r)) and abs(num(C("bancaria", r)) - 5100) < 0.01]
            if len(hits5100) > 1:
                add("ALTA", f"{ws.title}!col-bancária({mapa['bancaria']})", "T12-taxa-bancaria",
                    f"Taxa 5.100 imputada {len(hits5100)}× (linhas {[r for r, _ in hits5100][:10]}…).",
                    "Fórmula arrastada sem referência absoluta.", round((len(hits5100)-1)*5100, 2),
                    "Centralizar 5.100 numa célula (ex.: PRESSUPOSTOS!$B$10); ratear por processo, não replicar.", "planilha")

        # --- C5: HS x taxa + banda válida
        if "hs" in mapa and "taxa" in mapa:
            for r in range(hr + 1, nrows + 1):
                if C("desc", r) is None and C("item", r) is None: continue
                hsraw, tx = str(C("hs", r) or "").replace(".", "").replace(" ", ""), pct(C("taxa", r))
                if tx is None: continue
                if tx not in BANDAS:
                    add("ALTA", L("taxa", r), "C5-banda-invalida",
                        f"L{r}: taxa {tx}% fora das bandas MFN (0/2.5/5/7.5/20).",
                        "Digitado à mão ou banda inexistente.", "modelo",
                        "Corrigir para a banda pautal correcta (validar c/ despachante).", "item")
                fam = hsraw[:4]
                if fam in HS_REF:
                    ref, cat, sub, conf = HS_REF[fam]
                    if abs(tx - ref) > 0.01 and not (fam == "8419" and tx == 0.0):
                        cif = num(C("cif", r)) or 0.0
                        add("ALTA", L("taxa", r), "C5-hs-vs-taxa",
                            f"L{r}: HS {hsraw} ({sub}) taxado a {tx}% vs referência {ref}% [confiança: {conf}].",
                            "Taxa única aplicada a toda a planilha ou HS errado.", round(abs(tx-ref)/100*cif, 2),
                            f"Adoptar {ref}% (ou justificar HS alternativo com despachante).", "item")
                # categoria x HS
                if "categoria" in mapa:
                    cat = norm(C("categoria", r))
                    famcat = {"3822": "reagente", "3002": "reagente", "3926": "consum", "7017": "consum",
                              "4016": "consum", "7326": "consum", "7323": "consum"}.get(fam, "equip")
                    if fam == "7323" and "equip" in cat:
                        add("ALTA", L("categoria", r), "C9-cat-vs-hs",
                            f"L{r}: categoria '{C('categoria', r)}' incompatível com HS {hsraw} (utensílio doméstico → 20%).",
                            "Item de cozinha classificado como equipamento lab.", "modelo",
                            "Obter datasheet 'laboratory sterilizer' e reclassificar 8419.20, ou assumir 7323.93 = 20%.", "item")
                    elif famcat == "consum" and "equip" in cat and fam not in ("7323",):
                        add("MEDIA", L("categoria", r), "C9-cat-vs-hs",
                            f"L{r}: HS {hsraw} é de consumível mas categoria = '{C('categoria', r)}'.",
                            "Categoria genérica 'Equipamento' em consumíveis.", "modelo",
                            "Alinhar categoria à família HS (ver HS_REF).", "item")

        # --- C6: IVA 16% sobre base cumulativa
        if "iva" in mapa and "cif" in mapa:
            for r in range(hr + 1, nrows + 1):
                if C("desc", r) is None and C("item", r) is None: continue
                iva, cif = num(C("iva", r)), num(C("cif", r))
                d = num(C("direitos", r)) if "direitos" in mapa else 0.0
                if iva is not None and cif is not None and cif > 0:
                    esp = round((cif + (d or 0.0)) * 0.16, 2)
                    if abs(esp - iva) > max(TOL, esp * 0.01):
                        add("MEDIA", L("iva", r), "C6-iva-base",
                            f"L{r}: IVA {iva:,.2f} ≠ 16%×(CIF+direitos)={esp:,.2f}.",
                            "IVA sobre base errada (sem direitos / sobre preço de venda).", round(esp-iva, 2),
                            f"=ROUND(({Cf('cif', r)}+{Cf('direitos', r)})*16%;2) → {esp:,.2f}.", "item")

        # --- C8: markup implícito (dispersão)
        if "markup" in mapa:
            mks = [(pct(C("markup", r)), r) for r in range(hr + 1, nrows + 1)
                   if pct(C("markup", r)) is not None and (C("desc", r) is not None or C("item", r) is not None)]
            if mks:
                vals = Counter(round(m, 2) for m, _ in mks)
                moda = vals.most_common(1)[0][0]
                for m, r in mks:
                    if abs(m - moda) > 0.01:
                        add("MEDIA", L("markup", r), "T6-markup-divergente",
                            f"L{r}: markup {m:.2f}% ≠ padrão {moda:.2f}%.",
                            "Markup diferenciado sem justificação ou erro de digitação.", "modelo",
                            f"Uniformizar a {moda:.2f}% ou justificar categoria (reagente/consumível/equipamento).", "item")

        # --- C7: bloco despacho/totais (10% taxas? 15.22% impacto?) + reconciliação SUM
        for r in range(hr + 1, nrows + 1):
            lbl = (norm(C("desc", r)) + " " + norm(C("item", r))).strip()
            if not lbl: continue
            if "total" in lbl or "despacho" in lbl or "desembaraco" in lbl or "desalfandeg" in lbl:
                if "total" in lbl:
                    for k in ("direitos", "iva", "despacho", "total", "cif", "frete", "seguro"):
                        if k not in mapa: continue
                        v = num(C(k, r))
                        if v is None: continue
                        s = sum(num(C(k, rr)) or 0.0 for rr in range(hr + 1, r))
                        if abs(s - v) > max(1.0, abs(v) * 0.001):
                            add("ALTA", L(k, r), "C7-total-sum",
                                f"L{r} ({lbl[:40]}): {k}={v:,.2f} ≠ SUM={s:,.2f} (dif. {v-s:,.2f}).",
                                "Linha excluída do SUM ou valor manual.", round(v-s, 2),
                                f"=SUM({Cf(k, hr+1)}:{Cf(k, r-1)}) → {s:,.2f}.", "planilha")
                for k in ("taxa", "markup"):
                    if k in mapa and num(C(k, r)) is not None and ("despacho" in lbl or "desembaraco" in lbl or "taxa" in lbl):
                        add("INFO", L(k, r), "C7-bloco-despacho",
                            f"L{r} ({lbl[:40]}): {k}={pct(C(k, r)):.2f}% — validar manualmente.",
                            "Percentual de bloco (ex.: 10% taxas / 15.22% desembaraço).", "n/d",
                            "Confirmar base do % e recalcular contra o SUM do bloco.", "planilha")

        # --- C10: modo de frete x origem + frete ausente
        if "modofrete" in mapa and "origem" in mapa:
            for r in range(hr + 1, nrows + 1):
                if C("desc", r) is None and C("item", r) is None: continue
                ori, modo = norm(C("origem", r)), norm(C("modofrete", r))
                if not modo:
                    add("MEDIA", L("modofrete", r), "C10-modo-ausente",
                        f"L{r}: modo de frete vazio (origem {C('origem', r)}).",
                        "Coluna não preenchida.", "modelo",
                        "Preencher: RSA→rodoviário; China/UE/Japão→marítimo (aéreo só urgentes).", "item")
                elif (("africa" in ori and "sul" in ori) or ori in ("rsa", "za", "r.s.a.")) and "rodov" not in modo:
                    add("MEDIA", L("modofrete", r), "C10-modo-origem",
                        f"L{r}: origem RSA com modo '{C('modofrete', r)}' (esperado rodoviário).",
                        "Modo incoerente com origem.", "modelo", "Corrigir modo e recalcular frete/CIF.", "item")
        if "frete" in mapa and "cif" in mapa:
            for r in range(hr + 1, nrows + 1):
                if C("desc", r) is None and C("item", r) is None: continue
                if (num(C("cif", r)) or 0) > 0 and num(C("frete", r)) is None:
                    add("MEDIA", L("frete", r), "C10-frete-zero",
                        f"L{r}: CIF sem frete lançado.", "Rateio não chegou ao item.", "modelo",
                        "Ratear frete do processo por peso-volume.", "item")

        # --- Células vermelhas (preenchimento directo)
        for r in range(hr + 1, nrows + 1):
            for k, c in mapa.items():
                cell = wsv.cell(r, c)
                if is_red_fill(cell):
                    add("ALTA", f"{ws.title}!{cell.coordinate}", "CELULA-VERMELHA",
                        f"L{r} col '{k}' com preenchimento vermelho (valor: {cell.value}).",
                        "Flag manual ou resultado de regra — validar causa.", "n/d",
                        " inspeccionar fórmula/origem; corrigir e limpar o flag.", "item")

    # --- Condicionais vermelhas do workbook
    for aba, rng, frm in condicoes_vermelhas(wb):
        add("INFO", f"{aba}!{rng}", "CONDICIONAL-VERMELHA",
            f"Regra condicional vermelha activa em {rng} ({frm}).",
            "Validação do autor da planilha.", "n/d",
            "Listar células actualmente vermelhas (ver CELULA-VERMELHA) e tratar uma a uma.", "planilha")

    # --- FASE 4: relatório
    ach.sort(key=lambda a: ({"ALTA": 0, "MEDIA": 1, "INFO": 2}[a.sev],
                            -(a.impacto_mt if isinstance(a.impacto_mt, float) else -1)))
    for i, a in enumerate(ach, 1): a.seq = i
    print(f"\n{'='*78}\nACHADOS: {len(ach)} "
          f"(ALTA={sum(1 for a in ach if a.sev=='ALTA')}, "
          f"MEDIA={sum(1 for a in ach if a.sev=='MEDIA')}, "
          f"INFO={sum(1 for a in ach if a.sev=='INFO')})")
    tot = sum(a.impacto_mt for a in ach if isinstance(a.impacto_mt, float) and a.impacto_mt > 0)
    print(f"Impacto quantificado (soma das diferenças positivas): {tot:,.2f} MT\n")
    for a in ach:
        imp = f"{a.impacto_mt:,.2f} MT" if isinstance(a.impacto_mt, float) else a.impacto_mt
        print(f"[{a.sev}] #{a.seq} {a.local} | {a.teste}\n  Achado: {a.achado}\n  Causa: {a.causa}\n"
              f"  Impacto: {imp} | Âmbito: {a.ambito}\n  Correcção: {a.correcao}\n")
    if out_csv:
        with open(out_csv, "w", newline="", encoding="utf-8") as fh:
            w = csv.writer(fh)
            w.writerow(["seq", "severidade", "localizacao", "teste", "achado", "causa_provavel",
                        "impacto_MT", "correcao", "ambito"])
            for a in ach: w.writerow(a.row())
        print(f"CSV gravado: {out_csv}")
    return ach

# ---------------------------------------------------------------- self-test (workbook sintético, erros plantados)
def self_test():
    from openpyxl import Workbook
    from openpyxl.styles import PatternFill
    from openpyxl.formatting.rule import CellIsRule
    wb = Workbook(); ws = wb.active; ws.title = "Orcamento"
    H = ["Item", "Descrição", "QTD", "Moeda", "FOB", "Câmbio", "FOB_MZN", "Frete", "Seguro",
         "CIF", "HS", "Taxa", "Direitos", "IVA", "Despacho", "Markup", "PU", "Categoria", "Origem", "Taxa Bancária"]
    ws.append(["PRESSUPOSTOS"]); ws.append(["USD", 64.00]); ws.append(["EUR", 69.50]); ws.append([])
    ws.append(H)
    rows = [
        # item, desc, qtd, moeda, fob, cambio, fob_mzn, frete, seguro, cif, hs, taxa, direitos, iva, desp, mk, pu, cat, origem, banc
        [1, "Centrífuga Z36", 2, "USD", 10000, 64.00, 640000.00, 20000, 13200.00, 673200.00, "8421.19", 0.05, 33660.00, 113097.60, 5000, 0.25, 900000, "Equipamento", "Alemanha", None],  # OK
        [2, "Reagente ELISA frio", 5, "USD", 2000, 63.00, 126000.00, 8000, 2680.00, 136680.00, "3822.00", 0.05, 6834.00, 22962.24, 1000, 0.25, 40000, "Reagente", "Alemanha", None],  # ERROS: câmbio 63≠64; taxa 5% vs 0%
        [3, "Tripés metal", 3, "EUR", 50, 69.50, 3475.00, 500, 79.50, 4054.50, "7326.90", 0.05, 202.73, 681.16, 200, 0.40, 4500, "Equipamento", "China", None],  # ERROS: taxa 5 vs 20; markup 40 vs 25; categoria
        [4, "Agitador IKA", 1, "USD", 1500, 64.00, 96000.00, 3000, 1980.00, 100980.00, "8479.82", 0.05, 5049.00, 16964.64, 800, 0.25, 150000, "Equipamento", "Alemanha", 5100],  # banc 1ª vez
        [5, "Barras PTFE", 5, "USD", 100, 64.00, 6400.00, 500, 100.00, 7000.00, "8479.90", 0.05, 350.00, 1080.00, 600, 0.25, 6800, "Consumível", "China", 5100],  # ERROS: seguro 100 vs 138; banc 5100 2ª vez
    ]
    for r in rows: ws.append(r)
    ws["I10"].fill = PatternFill("solid", fgColor="FFFF0000")  # flag vermelho manual (seguro item 5)
    ws.conditional_formatting.add("P6:P20", CellIsRule(operator="greaterThan", formula=["0.30"],
        fill=PatternFill("solid", fgColor="FFFFC7CE")))  # regra vermelha markup>30%
    path = "/tmp/selftest_orcamento.xlsx"
    wb.save(path)
    print("Workbook sintético gravado (5 itens, 7+ erros plantados). A auditar…")
    ach = auditar(path)
    print(f"SELF-TEST: {len(ach)} achados.")
    esperados = ["T1-cambio-unico", "C5-hs-vs-taxa", "T6-markup-divergente", "T13-seguro-2pct",
                 "T12-taxa-bancaria", "CELULA-VERMELHA", "CONDICIONAL-VERMELHA", "C9-cat-vs-hs"]
    testes = {a.teste for a in ach}
    falta = [t for t in esperados if t not in testes]
    print("Controlos disparados:", sorted(testes))
    print("EM FALTA:", falta if falta else "nenhum — SELF-TEST OK ✔")
    return not falta

if __name__ == "__main__":
    if len(sys.argv) == 2 and sys.argv[1] == "--self-test":
        sys.exit(0 if self_test() else 1)
    if len(sys.argv) < 2:
        print(__doc__); sys.exit(2)
    out = None
    if "--out" in sys.argv:
        out = sys.argv[sys.argv.index("--out") + 1]
    auditar(sys.argv[1], out)