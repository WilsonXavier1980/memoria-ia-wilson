# RELATÓRIO DE AUDITORIA FORENSE — Orçamentos IIAM / Infinity Health, SA

**Data:** 11/09/2026 · **Auditor:** Arena.ai (Agent Mode) · **Ref.:** Concursos 35A003641 (CP/03/2026, CL/03/2026, PD/05/2026)
**Branch:** `arena/01a091da-memoria-ia-wilson`

---

## 0. RESUMO EXECUTIVO — LEIA PRIMEIRO

### 0.1 O ficheiro pedido NÃO existe no repositório indicado

O URL indicado (`https://github.com/WilsonXavier1980/memoria-ia`) **não resolve** (repositório inexistente ou sem acesso).
O checkout disponível (`memoria-ia-wilson`, branch `arena/01a091da-memoria-ia-wilson`, commit `d82edbc`) contém **apenas 2 ficheiros**:

| # | Ficheiro | Páginas | Conteúdo |
|---|----------|---------|----------|
| A | `Modelo de Proposta Financeira Final Para Concursos.pdf` | 17 | Proposta equipamento de FRIO — corpo refere CL/03/2026, **capa refere PD/05/2026 Reagentes** (ver E02) |
| B | `PROPOSTA_IIAM_35A003641-CP-03-2026_VFINAL.pdf` | 15 | Proposta equipamento de LABORATÓRIO — CP/03/2026, 19 itens, total 25.775.780,00 MT |

**Consequência directa:** os controlos das Secções 1–2 do mandato (Markup %, câmbios USD/GBP/ZAR/CNY/EUR→MZN,
fretes aéreo/marítimo/rodoviário, Taxa Bancária 5.100 MZN, Seguro 2%, honorários de despachante, formatações
condicionais vermelhas, célula/linha/coluna) **não são auditáveis sem o workbook-fonte** (`Orcamento cliente.docx/xlsx`).
Nenhuma destas variáveis aparece nas propostas finais (que são documentos cliente, só com preço unitário/total + IVA).
Isto está registado como **E01 (P0)** e, em compensação, a Secção 1–2 abaixo entrega: (i) engenharia reversa do que é
inferível a partir dos PDFs, e (ii) um **protocolo fechado de verificação** com fórmulas exactas para correr no
ficheiro-fonte assim que for anexado.

### 0.2 O que FOI auditado linha-a-linha (100% do que existe)

- Planilha de preços VFINAL, 19/19 itens: quantidades × preço unitário, IVA 16% linha-a-linha, somatórios — **aritmeticamente correcta**.
- Planilha de preços FRIO (pág. 16, imagem), 6 rubricas: idem — **correcta excepto 0,01 MT de arredondamento (E13)**.
- Coerência interna (nº concurso, contactos, datas, garantias, origens, referências) — **14 inconsistências encontradas (E02–E15)**.
- Classificação HS + taxa pautal + IVA + categoria para **todos os 25 itens** (Secções 3–4) — parecer técnico independente.

### 0.3 Top-5 riscos (ordenados por impacto)

| Prior. | ID | Achado | Impacto |
|--------|----|--------|---------|
| P0 | E01 | Workbook-fonte ausente — markup/câmbio/frete/seguro inauditáveis | Bloqueador total das Secções 1–2 |
| P0 | E02 | Capa trocada no Modelo (PD/05/2026 Reagentes × corpo CL/03/2026 Frio) | Risco de desclassificação |
| P0 | E03 | 7/19 referências + marcas "[A CONFIRMAR]" na VFINAL; §5.3 contradiz a planilha | Risco de desclassificação (Caderno §8.2/15.1.2) |
| P0 | E04 | Item 4 FRIO ("câmara fria") usa a mesma geleira 450 L do item 1 — provável não-conformidade técnica | Risco de desclassificação técnica |
| P1 | E06 | 6 itens de alto risco pautal (taxa 20% vs 5%: frio 8418, AC 8415, tripés 7326, panelas 7323) — cada 15 pp de erro ≈ 15% do CIF (17,4% se IVA embutido no custo) | Maior risco financeiro |

Enquadramento pautal de referência: bandas MFN 0% · 2,5% · 5% · 7,5% · 20% sobre o valor CIF; IVA 16% sobre
(CIF + direitos + ICE) [2](https://deepbeez.com/import-duty/import-duties-mozambique) [3](https://taxsummaries.pwc.com/mozambique/corporate/other-taxes);
equipamento laboratorial/centrífugas a 5%, reagentes de diagnóstico e medicamentos a 0% [1](https://www.tralac.org/documents/publications/trade-data-analysis/2926-mozambique-intra-africa-trade-and-tariff-profile-june-2019/file.html);
equipamento médico 0%, máquinas industriais 2,5–5%, bens de consumo 20% [4](https://kabum.digital/tabela-direitos-aduaneiros-mocambique-2026/).

---

## 1. VALIDAÇÃO DE VARIÁVEIS DINÂMICAS (Markup / Câmbio / Quantidades)

### 1.1 Quantidades — VERIFICADO contra as planilhas finais

**VFINAL (CP/03/2026), págs. 12–13** — todas as quantidades reproduzidas e validadas:

| Item | Ref.ª | QTD | Unid. | PU (MT) | Total s/IVA (MT) | Estado |
|------|-------|-----|-------|---------|------------------|--------|
| 1 | BSC-1500IIB2-X | 6 | Un | 980.000,00 | 5.880.000,00 | ✔ |
| 2 | Z 36 HK | 2 | Un | 1.450.000,00 | 2.900.000,00 | ✔ |
| 3 | 5430 R | 1 | Un | 620.000,00 | 620.000,00 | ✔ |
| 4 | SX-50 SpeedFill | 2 | Un | 1.100.000,00 | 2.200.000,00 | ✔ |
| 5 | [A CONFIRMAR] | 2 | Un | 850.000,00 | 1.700.000,00 | ✔ aritm.; ✘ ref. E03 |
| 6 | [A CONFIRMAR] | 1 | Un | 1.250.000,00 | 1.250.000,00 | ✔ aritm.; ✘ ref. E03 |
| 7 | [A CONFIRMAR] | 1 | Un | 1.650.000,00 | 1.650.000,00 | ✔ aritm.; ✘ ref. E03 |
| 8 | [A CONFIRMAR] | 1 | Un | 1.500.000,00 | 1.500.000,00 | ✔ aritm.; ✘ ref. E03 |
| 9 | Tripé galvanizado 20 cm | 3 | Un | 4.500,00 | 13.500,00 | ✔ aritm.; ⚠ preço-chão E12 |
| 10 | C-MAG HS 10 | 2 | Un | 95.000,00 | 190.000,00 | ✔ |
| 11 | C-MAG HS 7 | 1 | Un | 68.000,00 | 68.000,00 | ✔ |
| 12 | Barras 50–60 mm (cx.10) | 5 | Cx | 6.800,00 | 34.000,00 | ✔ aritm.; ⚠ E12 |
| 13 | MCO-170AICUVH-PE | 1 | Un | 780.000,00 | 780.000,00 | ✔ |
| 14 | [A CONFIRMAR] | 1 | Un | 210.000,00 | 210.000,00 | ✔ aritm.; ✘ ref. E03 |
| 15 | Haematokrit 200 | 1 | Un | 180.000,00 | 180.000,00 | ✔ |
| 16 | [A CONFIRMAR] | 1 | Un | 2.400.000,00 | 2.400.000,00 | ✔ aritm.; ✘ ref. E03 |
| 17 | HVE-50 | 1 | Un | 420.000,00 | 420.000,00 | ✔ |
| 18 | MS 3 digital | 1 | Un | 55.000,00 | 55.000,00 | ✔ |
| 19 | [A CONFIRMAR] | 2 | Un | 85.000,00 | 170.000,00 | ✔ aritm.; ✘ ref. E03 |
| **Total** | | | | | **22.220.500,00** s/IVA · **3.555.280,00** IVA · **25.775.780,00** c/IVA | ✔ soma exacta, IVA = 16,0000% |

**FRIO (CL/03/2026), pág. 16 (imagem)** — validado por leitura da imagem:

| Item | Ref.ª | QTD | PU (MT) | Total s/IVA | IVA 16% | Total c/IVA | Estado |
|------|-------|-----|---------|-------------|---------|-------------|--------|
| 1 | HYCD-469A | 2 | 380.443,04 | 760.886,08 | 121.741,77 | 882.627,85 | ✔ |
| 2 | HYCD-282A | 3 | 261.209,99 | 783.629,97 | 125.380,80 | 909.010,77 | ✔ |
| 3/3.1–3.3 | PH1000+CALIB-01+A2000-EX+CALIB-02 | 1 | 51.137,76 | 51.137,76 | 8.182,04 | 59.319,80 | ✔ aritm.; ✘ estrutura E10 |
| 4 | HYCD-469A | 1 | 380.443,04 | 380.443,04 | 60.870,89 | 441.313,93 | ✔ aritm.; ✘ técnico E04 |
| 5 | HDG-86Z04-12 | 1 | 454.877,55 | 454.877,55 | 72.780,41 | 527.657,96 | ✔ aritm.; ⚠ sem unidade E12 |
| 6 | HISENSE 1 | 9 | 47.508,04 | 427.572,36 | 68.411,58 | 495.983,94 | ✔ aritm.; ✘ origem E05 |
| **Total declarado** | | | | **2.858.546,76** | **457.367,48** | **3.315.914,24** | ⚠ E13: soma das linhas = 457.367,49 / 3.315.914,25 (**+0,01 MT**) |

> **E13 — causa:** o total IVA foi calculado como 16% sobre o total s/IVA (2.858.546,76 × 0,16 = 457.367,4816 → 457.367,48),
> enquanto a soma dos IVAs das linhas já arredondados dá 457.367,49. Impacto financeiro: **0,01 MT (nulo)**;
> impacto formal: método inconsistente. **Correcção:** recalcular totais por soma das linhas (457.367,49 / 3.315.914,25)
> ou parametrizar todas as células com `ROUND(x;2)` e totais por `SUM`.

### 1.2 Markup (%) e Câmbio (USD/GBP/ZAR/CNY/EUR→MZN) — NÃO AUDITÁVEL (E01), protocolo + inferências

**E01 — Bloqueador.** Nenhum dos dois PDFs contém: moeda-fonte, taxa de câmbio, data-fonte do câmbio, % markup,
base do markup, frete, seguro ou taxas. As propostas finais apresentam apenas PU/Total + IVA. Indícios indirectos:

1. **Preços VFINAL 100% redondos** (980.000, 1.450.000, 620.000…) → markup/arredondamento aplicado no fim; impossível
   reverter ao CIF sem o workbook. Risco: arredondamentos para cima em concurso "Menor Preço" custam competitividade.
2. **Preços FRIO com cêntimos** (380.443,04; 261.209,99; 47.508,04…) → saída directa de fórmula (FOB×câmbio×(1+…)),
   sem arredondamento comercial. Inconsistência metodológica entre as duas propostas (uma arredonda, outra não).
3. **Itens 9 e 12 VFINAL** (4.500 MT tripé; 6.800 MT/cx barras) estão 1–2 ordens de grandeza abaixo dos demais —
   típico de rateio de fixos (frete/despacho/bancário) que não chegou a estes itens, ou de preço sem landed cost.
   Verificar no workbook se estes itens absorveram quota de frete+despacho (E12).

**PROTOCOLO DE VERIFICAÇÃO (a correr no `Orcamento cliente` assim que anexado):**

| # | Teste | Fórmula / procedimento | Critério de aprovação |
|---|-------|------------------------|-----------------------|
| T1 | Câmbio por moeda único | Para cada moeda M, `CÂMBIO_M = MZN/FOB_M` deve ser idêntico em TODAS as linhas dessa moeda (tolerância = 0) | Desvio > 0 → erro de câmbio |
| T2 | Fonte e data do câmbio | Uma célula/aba "PRESSUPOSTOS" com fonte (BM vs banco comercial), data e spread; todas as fórmulas apontam para ela (sem valores hardcoded) | Qualquer taxa digitada à mão → chumbar |
| T3 | Spread bancário | Câmbio usado ≥ taxa média BM do dia; se = taxa média sem spread, o custo está subavaliado ~1–3% | Exigir evidência do spread |
| T4 | Base do markup documentada | `PU_sIVA = (CIF + Direitos + FreteInt + DespachoRateado + FixosRateados) × (1+Markup)` **com IVA de importação EXCLUÍDO** (é recuperável no regime normal; Incluir = dupla contagem, ver E07) | IVA dentro da base → erro crítico |
| T5 | Markup vs margem | Confirmar se a coluna é *markup sobre custo* ou *margem sobre preço*: `Margem = Markup/(1+Markup)`. Erro clássico: pedir 20% margem e aplicar 20% markup (margem real = 16,7%) | Declarar explicitamente |
| T6 | Consistência do markup | Mesmo % em todos os itens, ou tabela por categoria justificada; qualquer célula divergente deve estar comentada | Divergência muda = erro |
| T7 | Cêntimos vs arredondamento | Política única (ex.: ROUND para 2 casas no workbook; arredondamento comercial só na proposta final) | Mistura de políticas = erro |
| T8 | Reconciliação | `SUM(Total s/IVA)` do workbook = 22.220.500,00 (VFINAL) / 2.858.546,76 (FRIO) ao cêntimo | Diferença ≠ 0 → erro |

---

## 2. FRETES, DESALFANDEGAMENTO, TAXA BANCÁRIA, SEGURO, HONORÁRIOS

### 2.1 Fretes (aéreo/marítimo/rodoviário) — direccional por origem (inferido; a confirmar no workbook)

| Origem (conforme propostas) | Itens | Modo esperado | Prazo contratual | Avaliação |
|---|---|---|---|---|
| China (11 itens VFINAL + 4 FRIO) | 1,5,6,7,8,9,12,14,16,19 + frio 1,2,4,5 | **Marítimo** (contentor consolidado); aéreo só p/ itens urgentes/pequenos | máx. 90 dias / 60 dias úteis | ✔ compatível; frete aéreo p/ autoclave 1000 L (item 16) seria proibitivo — confirmar modo |
| Alemanha (5 itens) | 2,3,10,11,15,18 | Marítimo (Hamburgo→Maputo) ou aéreo (IKA/pequenos) | máx. 90 dias | ✔ |
| Reino Unido (1 item) | 4 (SX-50) | Marítimo/aéreo | máx. 90 dias | ✔; ⚠ pós-Brexit sem preferência SADC/UE (E15) |
| Japão (2 itens) | 13,17 | Marítimo | máx. 90 dias | ✔ |
| RSA (1 rubrica) | Frio item 3 | **Rodoviário** (Joanesburgo→Maputo) | 30 dias úteis | ✔; elegível a preferência SADC c/ certificado (E15) |
| "Moz" (1 rubrica) | Frio item 6 | N/A (stock local?) | 15 dias úteis | ✘ origem por esclarecer (E05) |

**Controlos a correr no workbook (T9–T11):**
- **T9 — Rateio de frete:** frete total por contentor/processo rateado por peso-volume (não por valor nem por igual);
  itens 9/12 (leves e baratos) devem absorver pouco frete mas NÃO zero. Frete zero em qualquer linha importada = erro.
- **T10 — Impacto no desalfandegamento:** como o valor aduaneiro = **CIF**, cada metical de frete/seguro altera direitos
  e IVA: `ΔCusto = ΔCIF × (1 + taxa_direito) × (1,16 se IVA não recuperável; 1,00 se recuperável)`. Frete lançado
  *depois* do cálculo dos direitos (fora do CIF) = erro que **subestima direitos+IVA**.
- **T11 — Coerência modal:** prazo 15 dias úteis (item 6) é incompatível com marítimo China→Maputo → corrobora stock
  local; se o workbook lhe imputa frete marítimo + direitos, há dupla imputação (E05).

### 2.2 Taxa Bancária 5.100 MZN — controlo T12

> **T12:** `COUNT` de imputações da taxa no workbook deve ser **= 1 por processo de importação/pagamento ao fornecedor**
> (não 1 por item). Sintoma clássico: `5.100 × 19 itens = 96.900 MZN` embutidos na VFINAL em vez de 5.100 MZN
> (sobrepreço ≈ 91.800 MZN ≈ 0,4% da proposta — pequeno mas revela erro de modelo que se repete noutros fixos).
> Nos PDFs não há vestígio da taxa (esperado); a verificação fica pendente do workbook.

### 2.3 Seguro 2% — controlo T13

> **T13 — Base correcta:** `Seguro = 2% × (FOB_MZN + Frete_MZN)` (construção do CIF; prática AT/Moçambique).
> **Erros típicos:** (a) 2% sobre o CIF final (circularidade — resolve-se por iteração ou pela fórmula
> `Seguro = 0,02/0,98 × (FOB+Frete)` se a apólice for "2% sobre CIF"); (b) 2% só sobre FOB (subestima);
> (c) % diferente por linha sem justificação. Impacto de (a)/(b): ±0,04% do CIF por item — sistemático,
> mas combinado com T10 propaga-se a direitos+IVA.

### 2.4 Honorários de despachante — controlo T14

> **T14:** honorários são por **DU/processo (+ ad-valorem eventual)**, nunca um % fixo por item. Verificar no workbook:
> (i) parcela fixa rateada pelo nº de linhas do processo; (ii) parcela variável (se existir) sobre o CIF e à taxa
> contratada; (iii) sem "IVA sobre honorários" embutido no custo se o IVA for recuperável (mesma lógica E07).
> Benchmark de sanidade: honorários totais tipicamente 0,5–2% do CIF para processos consolidados; fora disto, pedir factura pró-forma do despachante.

---

## 3. VERIFICAÇÃO DE HONORÁRIOS E TAXAS ADUANEIRAS POR ITEM (HS · Taxa · IVA · Categoria)

Metodologia: classificação indicativa pelo Sistema Harmonizado (SH 2022, que Moçambique adoptou na revisão pautal
vigente) [5](https://www.economiaemercado.com/artigo/mocambique-decide-finalmente-isentar-direitos-aduaneiros-na-importacao-de-veiculos-electricos),
taxa MFN de referência para Moçambique (bandas 0/2,5/5/7,5/20) [2](https://deepbeez.com/import-duty/import-duties-mozambique),
IVA-importação 16% sobre (CIF+direitos+ICE) [3](https://taxsummaries.pwc.com/mozambique/corporate/other-taxes).
**Recomendação transversal:** confirmar cada taxa na Pauta vigente junto do despachante/AT antes de fechar o workbook —
a classificação final é acto do despachante/AT, não desta auditoria.

### 3.1 VFINAL — 19 itens (CP/03/2026)

| Item | Descrição | HS proposto | Taxa MFN | IVA | Categoria | Notas / risco |
|------|-----------|-------------|----------|-----|-----------|---------------|
| 1 | Cabine biossegurança Cl. II | **8414.80** (capelas/hoods) | **5%** | 16% | Equip. laboratório | Alternativa 8419.89 (5%) — neutra |
| 2 | Centrífuga refrig. 500 mL | **8421.19** | **5%** | 16% | Equip. laboratório | Taxa confirmada p/ centrífugas lab [1](https://www.tralac.org/documents/publications/trade-data-analysis/2926-mozambique-intra-africa-trade-and-tariff-profile-june-2019/file.html) |
| 3 | Centrífuga refrig. 15 mL | **8421.19** | **5%** | 16% | Equip. laboratório | Idem |
| 4 | Enchedora SX-50 | **8422.30** | **5%** | 16% | Equip. lab/industrial | Máq. encher/fechar/rotular |
| 5 | Capsuladora | **8422.30** | **5%** | 16% | Equip. lab/industrial | Ref. pendente (E03) impede confirmação |
| 6 | Rotuladora automática | **8422.30** | **5%** | 16% | Equip. lab/industrial | Idem |
| 7 | Enchimento+enfrascamento | **8422.30** | **5%** | 16% | Equip. lab/industrial | Idem |
| 8 | Lavadora frascos/placas | **8422.19** (≈8422.20) | **5%** | 16% | Equip. laboratório | Ref. pendente (E03) |
| 9 | Tripés galvanizados 20 cm | **7326.90** ⚠ | **20%** (confirmar; possível 7,5%) | 16% | **Consumível metal** | **E06:** se o workbook usou 5% "equip. lab", subcusto ≈ 15% do CIF do item |
| 10 | Agitador C-MAG HS 10 | **8479.82** (≈8419.89) | **5%** | 16% | Equip. laboratório | Mistura/agitação |
| 11 | Agitador C-MAG HS 7 | **8479.82** | **5%** | 16% | Equip. laboratório | Idem |
| 12 | Barras magnéticas PTFE | **8479.90** (partes) ⚠ | **5%** (alt. 3926.90 = 20%) | 16% | **Consumível** | **E06:** defender 8479.90 como acessório de 8479.82; se AT classificar 3926.90 (plástico) → 20% |
| 13 | Incubadora CO₂ 180 L | **8419.89** | **5%** | 16% | Equip. laboratório | Tratamento térmico lab |
| 14 | Microscópio digital 1,3 MP | **9011.80** | **5%** (verificar 2,5% instr. científicos) | 16% | Equip. análise | Ref. pendente (E03) |
| 15 | Microhematócrito | **8421.19** | **5%** | 16% | Equip. laboratório | — |
| 16 | Autoclave 1.000 L | **8419.20** ⚠ | **0–5%** | 16% | Equip. lab/hospital | **Oportunidade:** fundamentar 0% como esterilizador médico [4](https://kabum.digital/tabela-direitos-aduaneiros-mocambique-2026/); item de 2,4 M MT — cada 5 pp ≈ 120.000 MT de CIF |
| 17 | Autoclave 50 L | **8419.20** ⚠ | **0–5%** | 16% | Equip. lab/hospital | Idem (420.000 MT) |
| 18 | Vortex MS 3 | **8479.82** | **5%** | 16% | Equip. laboratório | — |
| 19 | Panelas pressão digital 6 L | **8419.20** vs **7323.93** 🔴 | **0–5% vs 20%** | 16% | Equip. lab **ou** utensílio doméstico | **Maior risco unitário (E06):** sem ref./marca (E03) é impossível defender 8419.20; se AT ler "panela de pressão" → 7323.93 = 20%. Exigir datasheet a dizer "laboratory sterilizer/autoclave" |

### 3.2 FRIO — 6 rubricas (CL/03/2026)

| Item | Descrição | HS proposto | Taxa MFN | IVA | Categoria | Notas / risco |
|------|-----------|-------------|----------|-----|-----------|---------------|
| 1 | Combinado refrig.+freezer 450 L | **8418.10** 🔴 | **20%** | 16% | Equip. frio (bens consumo) | Bens de consumo = 20% [4](https://kabum.digital/tabela-direitos-aduaneiros-mocambique-2026/); 760.886 MT expostos |
| 2 | Combinado 282 L | **8418.10** 🔴 | **20%** | 16% | Equip. frio | 783.630 MT expostos |
| 3 | Termo-higrómetros + calibração | **9025.19** + serviço | **5%** (bens); **0%** (serviço) | 16% (bens; serviço cf. IVA interno) | Equip. medida + **serviço** | Separar calibração do valor aduaneiro (E10); origem RSA → preferência SADC c/ certificado (E15) |
| 4 | "Câmara fria" HYCD-469A | **8418.10** 🔴 | **20%** | 16% | Equip. frio | + **E04** (especificação) |
| 5 | Liofilizador −86 °C | **8419.39** | **5%** | 16% | Equip. laboratório | Secadores lab; over-spec vs −55/−80 °C do caderno — custo acrescido voluntário |
| 6 | AC Split 18.000 BTU ×9 | **8415.10** 🔴 | **20%** (se importado) | 16% | Equip. (consumo) | + **E05** (origem "Moz" — se stock local, sem direitos; se importado, 20%) |

### 3.3 Regra de ouro Moçambique (para o workbook)

> `Direitos = Taxa_HS × CIF` ; `IVA_import = 16% × (CIF + Direitos + ICE)` ; base do IVA é cumulativa [2](https://kabum.digital/simulador-direitos-aduaneiros-mocambique-2026/).
> Erro de 15 pp na taxa (5%→20%) = **+15% do CIF em direitos** (+2,4% do CIF em IVA não recuperável, i.e. **17,4%** se o IVA
> estiver embutido no custo — ver E07). Nos itens 8418/8415 (frio ≈ 2,35 M MT s/IVA), isto decide lucro/prejuízo.

---

## 4. ANÁLISE POR CATEGORIA DE PRODUTO

### 4.1 Reagentes (frio vs não-frio) — ⚠ CATEGORIA FANTASMA
**Não existe um único reagente em nenhuma das duas propostas** (ambas são 100% equipamento/consumíveis duráveis).
A capa do Modelo anuncia "REAGENTES E CONSUMÍVEIS DE LABORATÓRIO — PD/05/2026" (E02), mas o corpo é equipamento de frio.
Se existir um terceiro workbook/proposta (PD/05/2026) com reagentes, aplicar: reagentes de diagnóstico → **HS 3822.00/3002**,
taxa **0%** [1](https://www.tralac.org/documents/publications/trade-data-analysis/2926-mozambique-intra-africa-trade-and-tariff-profile-june-2019/file.html);
distinção frio/não-frio **não altera HS nem taxa** — altera apenas o **modo de frete** (cadeia de frio = aéreo +
embalagem validada) e o seguro. **Acção:** localizar o workbook PD/05/2026 ou corrigir a capa (E02).

### 4.2 Consumíveis — 3 itens, todos em risco pautal (E06)
- **Metal (VFINAL-9, tripés 7326.90):** taxa provável **20%** (artigos acabados de ferro/aço) — confirmar se workbook usou 5%.
- **Plástico/PTFE (VFINAL-12, barras):** defender **8479.90 a 5%** (partes de máquinas de mistura); risco de 3926.90 a 20%.
- **Vidro/borracha:** nenhum item nestas propostas — se existirem no workbook PD/05/2026, regra geral: vidro lab
  (7017.90) 5–7,5%, borracha (4016.90) 7,5–20% conforme acabamento; validar linha-a-linha.

### 4.3 Equipamentos (laboratório / hospital / análise) — núcleo das propostas, taxa-base 5%
- **Laboratório geral (8421, 8479, 8419.89, 8422):** 5% — 17 dos 19 itens VFINAL + liofilizador. Categoria alinhada com HS. ✔
- **Hospital/esterilização (8419.20, itens 16–17, ~2,82 M MT):** oportunidade de **0%** como equipamento médico —
  pedir ao despachante fundamentação escrita; ganho potencial ≈ 5% do CIF (≈ 140.000 MT s/IVA à escala da proposta).
- **Análise/medida (9011 microscópio, 9025 termo-higrómetro):** 5% (verificar 2,5% p/ instrumentos científicos na Pauta vigente).
- **Frio/Climatização (8418, 8415):** **20%** — excepção que confirma a regra; é aqui que mora o risco (E06).

### 4.4 Treinamento / Instalação / Calibração — tributação como SERVIÇO
- VFINAL §4/§6 e FRIO §5 incluem instalação, montagem, testagem, calibração inicial e formação **no preço dos bens**
  (correcto para o concurso: "preços incluem transporte, instalação, montagem e formação").
- Para o **custo de importação**, estes serviços **não integram o valor aduaneiro** (prestados após importação, em Moçambique)
  → **0% direitos**; IVA 16% aplica-se na factura final ao IIAM (já incluído), não no desembaraço.
- **E10 (frio item 3):** as 2 calibrações (CALIB-01/02) estão fundidas nos 51.137,76 MT sem decomposição → o workbook
  deve separar bens (9025, 5%, origem RSA) de serviços (0% direitos), sob pena de pagar direitos sobre serviços.
- **E07 — armadilha simétrica (verificar no workbook):** custos de serviço com IVA suportado em Moçambique: se a empresa
  está no regime normal de IVA (facturação/importação > 2,5 M MZN — é o caso) [3](https://taxsummaries.pwc.com/mozambique/corporate/other-taxes),
  esse IVA é **recuperável/creditável** e deve ficar **fora da base de markup**. Regra geral: **nenhum IVA recuperável
  entra no custo**.

---

## 5. GARGALOS E ERROS — REGISTO COMPLETO (localização exacta · causa · impacto)

| ID | Sev. | Localização exacta | Achado | Causa provável | Impacto no preço de venda |
|----|------|--------------------|--------|----------------|---------------------------|
| E01 | P0 | Repo inteiro: `Orcamento cliente.docx` ausente; URL `memoria-ia` inexistente | Workbook-fonte indisponível; markup/câmbio/frete/seguro/bancário/despachante/vermelhos inauditáveis | Ficheiro nunca subido ou repo errado indicado | **Bloqueador:** Secções 1–2 do mandato por executar |
| E02 | P0 | Modelo, **pág. 1 (capa, imagem)**: "PD/05/2026 — REAGENTES E CONSUMÍVEIS — 08/09/2026" vs corpo (págs. 8,10–11,15–16): "CL/03/2026 — equipamento de frio — 02/08/2026" | Capa de outra proposta encadernada neste PDF | Montagem/cópia de template | **Desclassificação** se submetido assim (objecto/concurso errados na capa) |
| E03 | P0 | VFINAL, **§5.3 págs. 9–10 + planilha págs. 12–13, itens 5,6,7,8,14,16,19**: Ref.ª/Marca "[A CONFIRMAR]"; §5.3 itens 1–4,9–13,15,17,18: refs válidas + sufixo "[A CONFIRMAR]" | 37% dos itens sem ref./marca/fabricante; contradição §5.3↔planilha | Cotação de fornecedor pendente à data de fecho (10/09/2026) | **Desclassificação** (Caderno §8.2 catálogos/origem + §15.1.2 autorização fabricante); sem datasheet, HS/taxa (E06) de 6,68 M MT s/IVA são indefensáveis |
| E04 | P0 | Modelo, **pág. 12 (Tabela 4, #04) + pág. 14 + pág. 16 item 4**: "Câmara Fria combinada" = HYCD-469A 450 L (idêntico ao item 1) | "Câmara fria" (normalmente walk-in, HS 8418.69/9406) respondida com armário 450 L | Interpretação do caderno ou erro de copy-paste da linha do item 1 | **Não-conformidade técnica** (441.314 MT c/IVA em risco) + HS/taxa errados se o bem correcto for outro |
| E05 | P0 | Modelo, **pág. 16, item 6, coluna "País de Origem"**: "Moz" (Hisense Split 18k BTU) | Origem de fabrico declarada como Moçambique; Hisense não fabrica AC em Moçambique | Confusão origem-fabrico vs fonte-aquisição (stock local) | Se importado: **direitos 20% + IVA s/ direitos omitidos ou mal lançados**; formal: coluna de origem falsa |
| E06 | P1 | **Taxas pautais (seis linhas):** VFINAL-9 (7326.90), VFINAL-12 (8479.90 vs 3926.90), VFINAL-19 (8419.20 vs 7323.93), FRIO-1/2/4 (8418.10), FRIO-6 (8415.10) | Risco de taxa 20% onde o modelo-padrão "equip. lab" sugere 5% | Aplicação de taxa única 5% a toda a planilha (a confirmar no workbook) | **Até +15% do CIF por item** (+17,4% se IVA embutido — E07). Ex.: frio 8418/8415 ≈ 2,35 M MT s/IVA: cada 5 pp ≈ 117.000 MT |
| E07 | P1 | Workbook (metodologia — a confirmar): base do markup | IVA de importação (16% s/ CIF+direitos) incluído na base de custo antes do markup | Modelo "custo total com IVA × markup" | **Dupla contagem:** IVA recuperável vira custo + margem + IVA outra vez na factura → preço **~10–14% acima** do necessário (fatal em "Menor Preço") ou margem fictícia |
| E08 | P1 | Workbook (a confirmar): coluna seguro | Base do seguro 2% mal definida (sobre CIF circular / só FOB / % variável) | Fórmula ambígua | ±0,04–0,1% do CIF por item, sistemático, propagado a direitos+IVA (T10) |
| E09 | P1 | Workbook (a confirmar): taxa bancária 5.100 | Taxa fixa lançada por item em vez de 1× por processo | Fórmula arrastada sem `$` / sem rateio | Até +91.800 MT na VFINAL (19×) se replicada; sintoma de erro de rateio de fixos |
| E10 | P1 | Modelo, **pág. 16, item 3/3.1–3.3**: 4 sub-linhas (2 aparelhos + 2 calibrações) fundidas em QTD 1 × 51.137,76 | QTD ambígua (1 conjunto? 1 de cada?); bens+serviços misturados | Montagem da tabela por merge de células | Direitos pagos sobre serviços (calibração) se valor não for segregado; quantidade pode não cumprir o mapa do caderno |
| E11 | P2 | Contactos: Modelo corpo **+258 848770340** (45×) vs VFINAL **+258 848670340** (17×) vs Modelo pág.16 **+258 86 226 0072**; `info@` vs `comercial@` | Três telefones distintos; dois e-mails | Templates de épocas distintas | Formal: dados do concorrente inconsistentes; risco de contacto falhado pelo júri |
| E12 | P2 | VFINAL **itens 9 (4.500 MT) e 12 (6.800 MT)**; Modelo pág. 16 **item 5 sem Unidade Física** | Preços-chão sem landed cost visível; célula de unidade vazia | Rateio de frete/despacho não chegou aos itens baratos; lapso de preenchimento | Margem negativa nos itens 9/12 se o rateio for corrigido sem repricing; item 5 = defeito formal |
| E13 | P2 | Modelo **pág. 16, linha Total**: IVA 457.367,48 vs soma-das-linhas 457.367,49; total 3.315.914,24 vs 3.315.914,25 | Total por 16%×Total em vez de SUM(linhas) | Método de totalização inconsistente | **0,01 MT** — nulo financeiramente; corrigir método (Secção 1.1) |
| E14 | P2 | VFINAL **pág. 11-e + pág. 14**: garantia "12 a 24 meses conforme fabricante" + garantia definitiva 2.000.000 MT (7,76% da proposta) | Garantia vaga; % da garantia definitiva por confirmar no caderno | Redacção padrão | Esclarecimento do júri; se caderno exigir valor/% diferente → não-conformidade |
| E15 | P2 | VFINAL item 4 (UK), 13/17 (Japão); Modelo item 3 (RSA) | Preferências comerciais: UK/JP = MFN sem preferência; RSA = preferência SADC só c/ certificado de origem | Falta de coluna "regime preferencial + certificado" | UK/JP: custo acrescido se workbook assumiu taxa UE/SADC; RSA: poupança perdida (5%→0%) se certificado não for junto |

*Nota sobre "formatações condicionais vermelhas": só existem no workbook-fonte (E01); nos PDFs não há células vermelhas
observáveis. O protocolo T1–T14 acima é o equivalente funcional dessas validações.*

---

## 6. SOLUÇÕES IMEDIATAS — PLANO DE CORRECÇÃO PRIORIZADO (impacto financeiro decrescente)

| Ordem | ID | Correcção específica | Âmbito | Esforço | Efeito |
|-------|----|----------------------|--------|---------|--------|
| 1 | E01 | **Anexar `Orcamento cliente` (xlsx/docx) + cotações de fornecedor + DU pró-forma do despachante** a este repo/branch; correr protocolo T1–T14 e reemitir este relatório como Rev.1 | Tudo | ⏳ depende do cliente | Desbloqueia auditoria real das Secções 1–2 |
| 2 | E06+E07 | No workbook: (a) criar coluna `HS_CODE` + `TAXA_PAUTAL` por item conforme tabelas §3.1/3.2 e validar com o despachante; (b) **expurgar todo o IVA recuperável da base de markup**; (c) recalcular PU; (d) fundamentar por escrito 8419.20→0% (autoclaves) e 8479.90 (barras) | 9 linhas + metodologia global | 1 dia + despachante | Decide lucro/prejuízo; até ~±15% do CIF nas linhas 20% + ~10–14% de competitividade (E07) |
| 3 | E03 | Obter refs/marcas/datasheets/autorizações p/ itens 5,6,7,8,14,16,19; remover TODOS os "[A CONFIRMAR]" (§5.3 e planilha); juntar catálogos+origens (Caderno §8.2/15.1.2) | 7 itens (6,68 M MT s/IVA) | Fornecedores | Elimina risco de desclassificação + fundamenta HS |
| 4 | E02 | Substituir a capa do Modelo pela correcta (CL/03/2026 — equipamento de frio) OU separar o processo PD/05/2026-reagentes em documento próprio com o seu workbook | Doc. Modelo | 1 h | Elimina risco de desclassificação |
| 5 | E04 | Confirmar no Caderno de Encargos o que é o item 04 ("câmara fria" walk-in vs armário); se walk-in, recotar (HS 8418.69/9406, taxa a confirmar) e republicar Tabela 4 + §5.3 + planilha | 1 linha (441 kMT) | Técnico + fornecedor | Elimina não-conformidade técnica |
| 6 | E05+E10 | Item 6: declarar origem de fabrico real (China/RSA…) ou mover "stock local" para coluna fonte distinta, zerando frete+direitos só se factura local; Item 3: desdobrar em 3.1 bens / 3.2 bens / 3.3–3.4 serviços, com QTD por sub-linha | 2 rubricas frio | 2 h | Direitos/IVA correctos; quantidades conformes o mapa |
| 7 | E09+E08+T12/T13 | Fixar no workbook: taxa bancária 5.100 em célula única por processo (referência absoluta); seguro `=2%×(FOB+Frete)` numa única fórmula nomeada; honorários por DU rateados por peso-volume | Metodologia global | 2 h | Elimina sobre/subcusto sistemático de fixos |
| 8 | E11+E12+E13+E14+E15 | Normalizar telefone/e-mail/NUIT em todos os docs; preencher unidade item 5 frio; repricing itens 9/12 após rateio correcto; alinhar totais (SUM linhas); fixar garantia "mínimo 12 meses" por item; confirmar % garantia definitiva no caderno; juntar certificado origem RSA / assumir MFN p/ UK-JP | Formal, 8 pontos | 3 h | Proposta limpa, sem pedidos de esclarecimento |

**Regra de paragem:** não submeter nenhuma das propostas antes de fechar os P0 (E01–E05). Os P1 (E06–E10) decidem
margem; os P2 (E11–E15) decidem imagem e pedidos de esclarecimento.

---

## ANEXO A — Fórmulas de referência do modelo de custeio (a implementar no workbook)

```
CIF_MZN      = FOB_moeda × CÂMBIO_M + FRETE_MZN + SEGURO_MZN
SEGURO_MZN   = 2% × (FOB_moeda × CÂMBIO_M + FRETE_MZN)
DIREITOS     = TAXA_HS × CIF_MZN
IVA_IMPORT   = 16% × (CIF_MZN + DIREITOS + ICE)        → NÃO entra no custo (recuperável, regime normal)
CUSTO_BASE   = CIF_MZN + DIREITOS + FRETE_INT + DESPACHO_rateado + FIXOS_rateados (bancário 5.100 1×/processo, etc.)
PU_sIVA      = CUSTO_BASE × (1 + MARKUP)               → MARKUP = markup sobre custo (declarar; ≠ margem)
TOTAL_LINHA  = ROUND(QTD × PU_sIVA; 2)
IVA_LINHA    = ROUND(TOTAL_LINHA × 16%; 2)
TOTAL_cIVA   = TOTAL_LINHA + IVA_LINHA
TOTAIS       = SUM das linhas (nunca % sobre o total — evita E13)
```

## ANEXO B — Rastreabilidade da auditoria
- `Modelo de Proposta Financeira Final Para Concursos.pdf` (17 págs; pág. 1 e 16 são imagem — validadas por leitura visual).
- `PROPOSTA_IIAM_35A003641-CP-03-2026_VFINAL.pdf` (15 págs; texto integral extraído e reconciliado).
- Cálculos de reconciliação executados por script (19/19 + 6/6 linhas) — resultados na Secção 1.1.
- Pauta/taxas: fontes citadas inline; validação final compete ao despachante/AT.

*Fim do relatório — Rev.0 (pendente de Rev.1 após recepção do workbook-fonte).*

---

## ANEXO C — Kit forense automatizado (Rev.0.1, 11/09/2026)
- `auditoria_planilha.py` — motor de auditoria (controlos T1/T2/T6/T12/T13/C2/C5/C6/C7/C9/C10 + células/regras vermelhas).
  Uso: `python3 auditoria_planilha.py "Planilha Orçamento V Preço de Venda.xlsx" --out achados.csv`
  Validação: `python3 auditoria_planilha.py --self-test` → **10/10 controlos OK ✔** (workbook sintético, erros plantados).
- `hs_reference_mz.csv` — tabela HS × taxa MFN × IVA × categoria (24 linhas) usada pelo controlo C5.
- Assim que o workbook-fonte for anexado, a Rev.1 deste relatório sai com achados célula-a-célula.
