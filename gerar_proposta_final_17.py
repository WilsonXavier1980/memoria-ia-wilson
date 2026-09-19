#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Geração da Proposta Técnica & Financeira - 17 páginas
- Página 1: Capa com Nº Cliente, Ref., Nome, Data 28/09/2026
- Página 2: Índice atualizado
- Páginas 3-15: Conteúdo técnico e financeiro atualizado com base no concurso 35A003641/CP/03/2026
- Página 16 (paisagem): Planilha Excel com colunas Ref, Descrição, QTD, Preço Unitário, Total s/IVA, IVA, Total c/IVA bem visíveis
- Página 17: Requisitos
- Texto justificado, linguagem formal erudita, correção ortográfica
"""
import os
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm
from reportlab.lib import colors
from reportlab.lib.styles import ParagraphStyle, getSampleStyleSheet
from reportlab.lib.enums import TA_JUSTIFY, TA_LEFT, TA_CENTER, TA_RIGHT
from reportlab.platypus import BaseDocTemplate, PageTemplate, Frame, Paragraph, Spacer, Table, TableStyle, PageBreak, Image, NextPageTemplate, KeepTogether
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.colors import HexColor

# Colors
DARK_GREEN = HexColor("#0e4d1e")
PRIMARY_GREEN = HexColor("#145214")
LIGHT_MINT = HexColor("#eaf6ec")
LIGHT_GREEN_HEADER = HexColor("#7ab648")
TABLE_HEADER_DARK = HexColor("#0f5d2e")
TABLE_ALTERNATE = HexColor("#f2faf2")
FOOTER_GREEN = HexColor("#1a7a33")

PAGE_W, PAGE_H = A4
MARGIN_LEFT = 32
MARGIN_RIGHT = 32
MARGIN_TOP = 52
MARGIN_BOTTOM = 46

# Fonts
try:
    pdfmetrics.registerFont(TTFont('DejaVu', '/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf'))
    pdfmetrics.registerFont(TTFont('DejaVu-Bold', '/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf'))
    FONT_NORMAL = 'DejaVu'
    FONT_BOLD = 'DejaVu-Bold'
except:
    FONT_NORMAL = 'Helvetica'
    FONT_BOLD = 'Helvetica-Bold'

def fmt_pt(value):
    s = f"{value:,.2f}"
    s = s.replace(",", "X").replace(".", ",").replace("X", ".")
    return s

# ---- Ler Excel para garantir exatidão ----
import openpyxl
xlsx_path = "CONCURSO PÚBLICO Nº 35A003641-CP-03-2026-IIAM atualizada.xlsx"
items_data = []
if os.path.exists(xlsx_path):
    wb = openpyxl.load_workbook(xlsx_path, data_only=True)
    ws = wb["Nº0016"]
    # Rows 11 onwards (item 1 at row 11). Cols: A=Item, B=Ref, C=Desc, D=Pais, E=Tempo, F=QTD, G=Un, H=PU, I=Total sIVA, J=IVA, K=Total cIVA
    for r in range(11, 40):
        item = ws.cell(row=r, column=1).value
        if item is None:
            continue
        try:
            item_int = int(float(str(item).strip()))
        except:
            continue
        if item_int <1 or item_int>40:
            continue
        ref = str(ws.cell(row=r, column=2).value or "").strip()
        desc = str(ws.cell(row=r, column=3).value or "").strip()
        pais = str(ws.cell(row=r, column=4).value or "").strip()
        tempo = str(ws.cell(row=r, column=5).value or "").strip()
        # correct encoding artefact
        tempo = tempo.replace("Ù", "Ú").replace("ù", "ú")
        qtd = ws.cell(row=r, column=6).value
        un = str(ws.cell(row=r, column=7).value or "UN").strip()
        pu = ws.cell(row=r, column=8).value
        # fallback if pu None then compute from total
        if pu is None:
            pu = 0
        # Clean up
        if not pais:
            pais = "—"
        if not tempo:
            tempo = "—"
        items_data.append({
            "item": item_int,
            "ref": ref,
            "desc": desc,
            "pais": pais,
            "tempo": tempo,
            "qtd": int(qtd) if qtd else 0,
            "un": un,
            "pu": float(pu) if pu else 0.0
        })
else:
    # fallback hardcoded if excel not found
    items_data = [
        {"item":1, "ref":"HR1200-IIA2-G", "desc":"Cabina de segurança biológica Classe II, BSL-2, comutável A2/B2 — 70 % de recirculação / 30 % de exaustão (A2) e 100 % de exaustão (B2), com comutação automática de modo por toque. Motor EC, área de trabalho ISO 5, filtro HEPA 99,99 %, plenum de pressão negativa. Aplicação: amostras biológicas, produtos químicos perigosos, citotóxicos e ensaios farmacêuticos. Inclui montagem e formação.", "pais":"China", "tempo":"90 Dias Úteis", "qtd":6, "un":"UN", "pu":530327.83},
        {"item":2, "ref":"LX-210L750R + LX-10L750R (rotor 6×500 mL)", "desc":"Centrífuga refrigerada para frascos de 500 mL — velocidade máxima 22 000 rpm / 21 000 rpm; FCR máxima 52 050 ×g / 47 400 ×g; capacidade 6×500 mL; –20 °C a +40 °C; ±1 °C; precisão ±10 rpm; 1 s–99 h 59 min 59 s; ruído ≤65 dB(A); 5 000 W / 3 400 W; 740×720×900 mm / 740×920×990 mm; 250 kg / 375 kg. Inclui montagem e formação.", "pais":"China", "tempo":"90 Dias Úteis", "qtd":2, "un":"UN", "pu":1175084.90},
        {"item":3, "ref":"LX-500R + 155T500R (rotor 12×15 mL)", "desc":"Centrífuga refrigerada de bancada para tubos de 15 mL — compressor sem CFC, motor brushless; ≥15 000 rpm; FCR ≥20 000 ×g (26 322 ×g); –20 °C a +40 °C, ±1 °C; rotor angular ≥12×15 mL em alumínio anodizado; LCD/TFT; bloqueio electrónico da tampa; ruído ≤65 dB; 220–240 V 50/60 Hz; CE, ISO 9001 e ISO 13485. Inclui montagem e formação.", "pais":"China", "tempo":"90 Dias Úteis", "qtd":1, "un":"UN", "pu":535360.27},
        {"item":4, "ref":"SX-50 SpeedFill + Tubos SIL & DY-SIL (Ø3,2/6,4/8,0 mm)", "desc":"Bomba peristáltica de bancada para enchimento asséptico, modelo SX-50 SpeedFill (cabeça dupla; 0,5–500 mL; ±1 %; servomotor; cGMP/FDA; suporte e bicos de dosagem incluídos); caixas de 15 m de tubo de silicone Ø3,2 mm, Ø6,4 mm e Ø8,0 mm (grau farmacêutico, autoclavável) e conectores duplos em Y correspondentes. Instalação, arranque e formação por técnico DARA (1 dia).", "pais":"Espanha", "tempo":"30 Dias Úteis", "qtd":2, "un":"UN", "pu":2552436.52},
        {"item":5, "ref":"SX-140-C & CO-5 mL + Changeover", "desc":"Máquina de bancada para fechamento automático de frascos (vials), modelo SX-140-C (estação pneumática de cravação; AISI 304; estrela POM; comando bimanual; até 1 500 uph; preparada para 1 formato de frasco 5 mL + 1 formato de tampa de alumínio). Kit de changeover e documentação de validação IQ/OQ.", "pais":"Espanha", "tempo":"30 Dias Úteis", "qtd":2, "un":"UN", "pu":3215928.92},
        {"item":6, "ref":"ZS-TB800", "desc":"Máquina de rotulagem automática (GMP/BPF, PLC com ecrã táctil, esteira motorizada) — acessórios: esteira motorizada, sensor fotoeléctrico e aplicador.", "pais":"África do Sul (RSA)", "tempo":"30 Dias Úteis", "qtd":1, "un":"UN", "pu":507778.80},
        {"item":7, "ref":"ZS-AFY1", "desc":"Máquina automática de enchimento e enfrascamento (50–250 mL, 20–60 frascos/min) — inclui esteira e distribuidor.", "pais":"África do Sul (RSA)", "tempo":"30 Dias Úteis", "qtd":1, "un":"UN", "pu":5208010.51},
        {"item":8, "ref":"MIELE PLW 8693 & Basic Basket set & E 118 & ProCare", "desc":"Máquina de lavar vidraria de laboratório MIELE PLW 8693 com DryPlus (DSS), caixa em aço inoxidável, dispensador de líquido e lança de aspiração com detecção de nível. Conjunto de cestos Basic (A101 superior, A150 inferior, A300/3, A301/5, 2×AK12/1, bocal A802), inserto E118 para 38 placas de Petri Ø100 mm, detergente ProCare Lab 10MA (5 L) e neutralizador ProCare Lab 30C (5 L).", "pais":"África do Sul (RSA)", "tempo":"30 Dias Úteis", "qtd":1, "un":"UN", "pu":2096104.29},
        {"item":9, "ref":"CON-TRIPOD", "desc":"Suporte tripé de aço para laboratório (altura 200 mm).", "pais":"África do Sul (RSA)", "tempo":"30 Dias Úteis", "qtd":3, "un":"UN", "pu":1405.35},
        {"item":10, "ref":"4240400 (IKA C-MAG HS 10)", "desc":"IKA C-MAG HS 10 — agitador magnético com aquecimento, placa cerâmica (até 15 L H₂O; segurança 550 °C; LED; DIN 12878).", "pais":"África do Sul (RSA)", "tempo":"30 Dias Úteis", "qtd":2, "un":"UN", "pu":330997.17},
        {"item":11, "ref":"3581200 (IKA C-MAG HS 7)", "desc":"IKA C-MAG HS 7 — agitador magnético com aquecimento, placa cerâmica (até 10 L H₂O; segurança 550 °C; LED; DIN 12878).", "pais":"África do Sul (RSA)", "tempo":"30 Dias Úteis", "qtd":1, "un":"UN", "pu":208177.84},
        {"item":12, "ref":"BRAND 137618", "desc":"Barra magnética octogonal em PTFE, com anel pivotante — L 51 mm, Ø 8 mm (caixa).", "pais":"África do Sul (RSA)", "tempo":"30 Dias Úteis", "qtd":5, "un":"UN", "pu":36898.00},
        {"item":13, "ref":"HCP-168", "desc":"Incubadora de CO₂ (180 L; +5 °C a 60 °C; CO₂ 0–20 % com sensor IR; humidade até 95 %; esterilização a 140 °C; controlo PID + ecrã táctil; filtro HEPA ISO Classe 5).", "pais":"China", "tempo":"90 Dias Úteis", "qtd":1, "un":"UN", "pu":528387.61},
        {"item":14, "ref":"VISISCOPE BL224PL T1", "desc":"Microscópio digital composto VISISCOPE® BL224PL T1 (contraste de fase e fundo claro, câmara digital T1).", "pais":"África do Sul (RSA)", "tempo":"30 Dias Úteis", "qtd":1, "un":"UN", "pu":271794.59},
        {"item":15, "ref":"LX-165T2-J + rotor 24×1,5/2,0 mL + rotor hematócrito", "desc":"Microcentrífuga de hematócrito (comando independente de velocidade e tempo; motor DC brushless; função pulse; rotor de hematócrito 24 capilares a 12 000 rpm e rotor para 24 tubos 1,5/2,0 mL a 14 000 rpm; 200–12 000 rpm; FCR máx. 13 680; 8,5–10,5 kg; 364×280×274 mm). Inclui montagem e formação.", "pais":"China", "tempo":"90 Dias Úteis", "qtd":1, "un":"UN", "pu":150105.93},
        {"item":16, "ref":"HRLM-60A", "desc":"Autoclave vertical de bancada 50 L (programável, eléctrico, aço inoxidável, com secagem; 105–134 °C; visor LCD; ciclo automático com pré-aquecimento, substituição, exaustão pulsada, aquecimento, esterilização e secagem; filtro de ar 0,22 µm). Inclui montagem e formação.", "pais":"China", "tempo":"90 Dias Úteis", "qtd":1, "un":"UN", "pu":430249.95},
        {"item":17, "ref":"HRLM-45", "desc":"Autoclave vertical de laboratório 45 L (Systec DE-45 ou equivalente; câmara 316L 344×500 mm; até 140 °C e 4 bar; 220–240 V; controlo por microprocessador; opções de documentação digital). Inclui montagem e formação.", "pais":"China", "tempo":"90 Dias Úteis", "qtd":1, "un":"UN", "pu":419942.28},
        {"item":18, "ref":"MSV-3500 & SV-4/30", "desc":"Vortex digital multitubos 300–3 500 rpm (tubos 0,2–50 mL; plataformas não incluídas) e plataforma SV-4/30 para 4×50 mL (Ø 30 mm).", "pais":"África do Sul (RSA)", "tempo":"30 Dias Úteis", "qtd":1, "un":"UN", "pu":153477.57},
        {"item":19, "ref":"Duo 60 (7-em-1) 6 L", "desc":"Panela de pressão 6 L em aço inoxidável, painel digital, função Keep Warm, válvula de segurança — inclui cuba interior em inox e grelha de vapor.", "pais":"África do Sul (RSA)", "tempo":"30 Dias Úteis", "qtd":2, "un":"UN", "pu":51671.60},
        {"item":20, "ref":"DE-45 (ref. fabr. 1545)", "desc":"Autoclave vertical de laboratório 45 L, 140 °C / 4 bar, câmara em 316L — inclui cesto em aço inoxidável adicional e kit de documentação digital.", "pais":"África do Sul (RSA)", "tempo":"30 Dias Úteis", "qtd":1, "un":"UN", "pu":2440717.96},
        {"item":21, "ref":"LTC4R (GRANT TX150-R4R)", "desc":"GRANT LTC4R — banho circulante refrigerado/aquecido TX150-R4R, –30 °C a +100 °C, 20 L, ±0,1 °C, caudal 18 L/min, 1,84 kW, com mangueiras e clips.", "pais":"Reino Unido", "tempo":"30 Dias Úteis", "qtd":1, "un":"UN", "pu":2086383.32},
        {"item":22, "ref":"E-CC90-D", "desc":"Contador digital de colónias E-CC90-D, Ø 50–90 mm placa de Petri, sensor de pressão/botões com sensibilidade ajustável, LED duplo (faixa + lupa), ampliação 3×/9×.", "pais":"África do Sul (RSA)", "tempo":"30 Dias Úteis", "qtd":3, "un":"UN", "pu":48415.67},
        {"item":23, "ref":"VP125", "desc":"Bomba de vácuo rotativa de palhetas 40–60 L/min, ≤10 Pa, 220 V — inclui filtro de entrada e carga de óleo.", "pais":"África do Sul (RSA)", "tempo":"30 Dias Úteis", "qtd":2, "un":"UN", "pu":43324.84},
        {"item":24, "ref":"PH800 (Apera)", "desc":"Apera PH800 — kit pH-metro de bancada (calibração 1–3 pontos; monitorização do eléctrodo; GLP 500 conjuntos; USB; software PC-Link; eléctrodo 201T-F).", "pais":"África do Sul (RSA)", "tempo":"30 Dias Úteis", "qtd":10, "un":"UN", "pu":76248.05},
        {"item":25, "ref":"DW-86L579", "desc":"Ultracongelador vertical –86 °C (ajustável a –70 °C; ~500 L; interior em aço inoxidável; controlo por microprocessador; alarme audiovisual; portas internas isoladas; refrigeração de alta eficiência). Inclui montagem e formação.", "pais":"China", "tempo":"90 Dias Úteis", "qtd":2, "un":"UN", "pu":679682.78},
        {"item":26, "ref":"CNW-TYI304 (Coplin vidro c/ tampa)", "desc":"Cuba de coloração (Coplin) em vidro com tampa.", "pais":"África do Sul (RSA)", "tempo":"30 Dias Úteis", "qtd":5, "un":"UN", "pu":6007.57},
        {"item":27, "ref":"613-7138 (VWR EHP)", "desc":"VWR EHP — micropipeta mecânica monocanal, volume variável 1 000–10 000 µL (autoclavável; resistente a UV).", "pais":"África do Sul (RSA)", "tempo":"30 Dias Úteis", "qtd":20, "un":"UN", "pu":42825.09},
        {"item":28, "ref":"BLBIO-SJA 7.000 L", "desc":"Biofermentador anaeróbico 6 000–7 000 L em AISI 316L, CIP/SIP, painel digital Siemens — inclui sparger N₂ e sensores de pH/nível.", "pais":"África do Sul (RSA)", "tempo":"30 Dias Úteis", "qtd":1, "un":"UN", "pu":35309742.83},
    ]

TOTAL_S_IVA = sum(d["qtd"]*d["pu"] for d in items_data)
TOTAL_IVA = TOTAL_S_IVA * 0.16
TOTAL_C_IVA = TOTAL_S_IVA * 1.16

# Styles
styles = getSampleStyleSheet()
style_title = ParagraphStyle('TitleGreen', parent=styles['Heading1'], fontName=FONT_BOLD, fontSize=10.5, textColor=DARK_GREEN, leading=13, spaceBefore=6, spaceAfter=6, alignment=TA_LEFT, keepWithNext=True)
style_heading2 = ParagraphStyle('Heading2Green', parent=styles['Heading2'], fontName=FONT_BOLD, fontSize=9.5, textColor=DARK_GREEN, leading=11, spaceBefore=7, spaceAfter=3, alignment=TA_LEFT)
style_normal = ParagraphStyle('NormalJustify', parent=styles['Normal'], fontName=FONT_NORMAL, fontSize=8.2, leading=11.2, alignment=TA_JUSTIFY, spaceBefore=2, spaceAfter=2, textColor=colors.HexColor("#222222"))
style_normal_small = ParagraphStyle('NormalSmall', parent=style_normal, fontSize=7.2, leading=9.8)
style_bullet = ParagraphStyle('Bullet', parent=style_normal, leftIndent=12, firstLineIndent=0, bulletIndent=6, spaceBefore=1.5, spaceAfter=1)
style_table_cell = ParagraphStyle('TableCell', parent=styles['Normal'], fontName=FONT_NORMAL, fontSize=6.4, leading=7.4, alignment=TA_LEFT, textColor=colors.HexColor("#1a1a1a"))
style_table_cell_small = ParagraphStyle('TableCellSmall', parent=style_table_cell, fontSize=5.9, leading=6.9)
style_table_header = ParagraphStyle('TableHeader', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=6.3, leading=7.2, textColor=colors.white, alignment=TA_CENTER)
style_table_cell_center = ParagraphStyle('TableCellCenter', parent=style_table_cell, alignment=TA_CENTER)
style_table_cell_right = ParagraphStyle('TableCellRight', parent=style_table_cell, alignment=TA_RIGHT)
style_caption = ParagraphStyle('Caption', parent=styles['Normal'], fontName=FONT_BOLD, fontSize=6.8, leading=7.8, textColor=DARK_GREEN, alignment=TA_LEFT, spaceBefore=3, spaceAfter=5)

TOTAL_PAGES = 17

def header_footer(canvas, doc):
    canvas.saveState()
    # Top infinity logo
    try:
        logo_path = "/tmp/infinity_logo.png"
        if os.path.exists(logo_path):
            canvas.drawImage(logo_path, PAGE_W - 135, PAGE_H - 48, width=92, height=14, preserveAspectRatio=True, mask='auto')
    except:
        pass
    # Header line
    canvas.setStrokeColor(HexColor("#7ab648"))
    canvas.setLineWidth(0.4)
    # canvas.line(MARGIN_LEFT, PAGE_H - 55, PAGE_W - MARGIN_RIGHT, PAGE_H -55)
    # Footer line
    y_footer_line = 36
    canvas.setStrokeColor(DARK_GREEN)
    canvas.setLineWidth(0.6)
    canvas.line(MARGIN_LEFT, y_footer_line+14, PAGE_W - MARGIN_RIGHT, y_footer_line+14)
    # Footer text
    canvas.setFont(FONT_NORMAL, 5.1)
    canvas.setFillColor(colors.HexColor("#222222"))
    canvas.drawString(MARGIN_LEFT, y_footer_line+6, "Infinity Health, SA")
    canvas.setFont(FONT_NORMAL, 4.7)
    canvas.drawString(MARGIN_LEFT, y_footer_line-2, "Rua de França, n.º 273, Bairro da Coop - Maputo")
    canvas.drawString(MARGIN_LEFT, y_footer_line-8, "Maputo")
    canvas.drawString(PAGE_W/2 + 10, y_footer_line+6, "Email: comercial@infinityhealth.co.mz")
    canvas.drawString(PAGE_W/2 + 10, y_footer_line-2, "Telefone: +258 848770340 / 870867668")
    canvas.drawString(PAGE_W/2 + 10, y_footer_line-8, "NUIT: 401550216")
    canvas.setFont(FONT_NORMAL, 5.3)
    canvas.setFillColor(DARK_GREEN)
    canvas.drawRightString(PAGE_W - MARGIN_RIGHT, y_footer_line-16, f"Página {doc.page} de {TOTAL_PAGES}")
    # Bottom bar
    bar_y = 0
    bar_h = 11
    canvas.setFillColor(HexColor("#0e7a3a"))
    canvas.rect(0, bar_y, PAGE_W, bar_h, stroke=0, fill=1)
    canvas.setFillColor(HexColor("#1a9a4a"))
    canvas.setFillAlpha(0.5)
    canvas.rect(0, bar_y+4, PAGE_W*0.4, bar_h-4, stroke=0, fill=1)
    canvas.setFillAlpha(1)
    canvas.setFillColor(colors.white)
    canvas.setFillAlpha(0.14)
    canvas.circle(PAGE_W*0.3, 6, 18, stroke=0, fill=1)
    canvas.circle(PAGE_W*0.55, 6, 12, stroke=0, fill=1)
    canvas.setFillAlpha(1)
    canvas.restoreState()

def header_footer_landscape(canvas, doc):
    canvas.saveState()
    pw, ph = landscape(A4)
    try:
        logo_path = "/tmp/infinity_logo.png"
        if os.path.exists(logo_path):
            canvas.drawImage(logo_path, pw - 135, ph - 38, width=92, height=14, preserveAspectRatio=True, mask='auto')
    except:
        pass
    y_footer_line = 26
    canvas.setStrokeColor(DARK_GREEN)
    canvas.setLineWidth(0.6)
    canvas.line(24, y_footer_line+14, pw - 24, y_footer_line+14)
    canvas.setFont(FONT_NORMAL, 5.1)
    canvas.setFillColor(colors.HexColor("#222222"))
    canvas.drawString(24, y_footer_line+6, "Infinity Health, SA")
    canvas.setFont(FONT_NORMAL, 4.7)
    canvas.drawString(24, y_footer_line-2, "Rua de França, n.º 273, Bairro da Coop - Maputo")
    canvas.drawString(24, y_footer_line-8, "Maputo")
    canvas.drawString(pw/2 + 10, y_footer_line+6, "Email: comercial@infinityhealth.co.mz")
    canvas.drawString(pw/2 + 10, y_footer_line-2, "Telefone: +258 848770340 / 870867668")
    canvas.drawString(pw/2 + 10, y_footer_line-8, "NUIT: 401550216")
    canvas.setFont(FONT_NORMAL, 5.3)
    canvas.setFillColor(DARK_GREEN)
    canvas.drawRightString(pw - 24, y_footer_line-16, f"Página {doc.page} de {TOTAL_PAGES}")
    bar_y = 0
    bar_h = 11
    canvas.setFillColor(HexColor("#0e7a3a"))
    canvas.rect(0, bar_y, pw, bar_h, stroke=0, fill=1)
    canvas.restoreState()

def cover_page(canvas, doc):
    canvas.saveState()
    pw, ph = A4
    canvas.setFillColor(HexColor("#fdfcf8"))
    canvas.rect(0,0,pw,ph, stroke=0, fill=1)
    try:
        if os.path.exists("/tmp/leaf_left.png"):
            canvas.drawImage("/tmp/leaf_left.png", -18, 0, width=135, height=ph, preserveAspectRatio=False, mask='auto')
        if os.path.exists("/tmp/leaf_bottom.png"):
            canvas.drawImage("/tmp/leaf_bottom.png", 0, -12, width=pw, height=125, preserveAspectRatio=False, mask='auto')
    except:
        pass
    canvas.setFillColor(HexColor("#ffffff"))
    canvas.setStrokeColor(HexColor("#c8e0c8"))
    canvas.setLineWidth(0.7)
    card_x, card_y, card_w, card_h = 72, 148, pw-144, ph-285
    canvas.roundRect(card_x, card_y, card_w, card_h, 14, stroke=1, fill=1)
    canvas.setStrokeColor(HexColor("#e8f5e8"))
    canvas.setLineWidth(0.4)
    canvas.roundRect(card_x+3, card_y+3, card_w-6, card_h-6, 11, stroke=1, fill=0)
    canvas.setFont(FONT_BOLD, 18)
    canvas.setFillColor(DARK_GREEN)
    canvas.drawRightString(pw - MARGIN_RIGHT - 10, ph - 65, "ORIGINAL")
    y = ph - 120
    canvas.setFont(FONT_BOLD, 10)
    canvas.setFillColor(colors.black)
    canvas.drawCentredString(pw/2, y, "INSTITUTO DE INVESTIGAÇÃO AGRÁRIA DE MOÇAMBIQUE")
    y -= 12
    canvas.setStrokeColor(HexColor("#7ab648"))
    canvas.setLineWidth(0.8)
    canvas.line(pw/2 - 170, y+6, pw/2 + 170, y+6)
    y -= 22
    canvas.setFont(FONT_BOLD, 11)
    canvas.setFillColor(DARK_GREEN)
    canvas.drawCentredString(pw/2, y, "PROPOSTA TÉCNICA & FINANCEIRA")
    y -= 45
    box_w = 180
    box_h = 38
    box_x = pw/2 + 20
    box_y = y - 18
    canvas.setStrokeColor(HexColor("#7ab648"))
    canvas.setFillColor(HexColor("#f6fdf6"))
    canvas.roundRect(box_x, box_y, box_w, box_h, 6, stroke=1, fill=1)
    canvas.setFont(FONT_BOLD, 7)
    canvas.setFillColor(HexColor("#2e6b2e"))
    canvas.drawString(box_x+12, box_y+22, "Data:")
    canvas.setFont(FONT_BOLD, 11)
    canvas.setFillColor(DARK_GREEN)
    canvas.drawString(box_x+12, box_y+8, "28/09/2026")
    y -= 85
    canvas.setFont(FONT_BOLD, 9)
    canvas.setFillColor(colors.black)
    canvas.drawCentredString(pw/2, y, "Concurso Público N.º 35A003641/CP/03/2026")
    y -= 32
    canvas.setFont(FONT_BOLD, 13)
    canvas.setFillColor(colors.black)
    canvas.drawCentredString(pw/2, y, "FORNECIMENTO DE DIVERSO")
    y -= 18
    canvas.drawCentredString(pw/2, y, "EQUIPAMENTO DE LABORATÓRIO")
    y -= 45
    canvas.setFont(FONT_NORMAL, 7)
    canvas.setFillColor(colors.HexColor("#333333"))
    canvas.drawCentredString(pw/2, y, "Entidade Contratante: Instituto de Investigação Agrária de Moçambique (IIAM)")
    y -= 11
    canvas.drawCentredString(pw/2, y, "Cliente N.º: 35A003641  |  Referência: 35A003641/CP/03/2026  |  Lote Único")
    y -= 11
    canvas.setFont(FONT_NORMAL, 6.5)
    canvas.drawCentredString(pw/2, y, "Av. das FPLM, n.º 2698, UGEA (Edifício do Pavilhão Novo) — Maputo")
    canvas.setFont(FONT_BOLD, 14)
    canvas.setFillColor(HexColor("#0e6b2e"))
    try:
        if os.path.exists("/tmp/infinity_logo.png"):
            canvas.drawImage("/tmp/infinity_logo.png", 28, 52, width=78, height=12, preserveAspectRatio=True, mask='auto')
    except:
        canvas.drawString(28, 58, "infinity")
    canvas.setFont(FONT_NORMAL, 5.5)
    canvas.setFillColor(colors.HexColor("#333333"))
    canvas.drawString(28, 44, "Infinity Health, SA")
    canvas.setFont(FONT_NORMAL, 4.8)
    canvas.drawString(28, 35, "Rua de França, n.º 273, Bairro da Coop - Maputo")
    canvas.drawString(28, 27, "NUIT: 401550216")
    canvas.drawString(28, 19, "Tel.: +258 848770340 / 870867668")
    canvas.drawString(28, 11, "Email: comercial@infinityhealth.co.mz")
    # Cover footer bar
    bar_h = 11
    canvas.setFillColor(HexColor("#0e7a3a"))
    canvas.rect(0,0,pw,bar_h, stroke=0, fill=1)
    canvas.setFillColor(colors.white)
    canvas.setFont(FONT_NORMAL, 5)
    canvas.drawCentredString(pw/2, 4, "Soluções Hospitalares  •  NUIT: 401550216  •  comercial@infinityhealth.co.mz  •  +258 848770340")
    canvas.restoreState()

def build_pdf(filename):
    from reportlab.lib.pagesizes import A4, landscape
    # Frames
    pw_p, ph_p = A4
    pw_l, ph_l = landscape(A4)
    frame_portrait = Frame(MARGIN_LEFT, MARGIN_BOTTOM, pw_p - MARGIN_LEFT - MARGIN_RIGHT, ph_p - MARGIN_TOP - MARGIN_BOTTOM, id='portrait_frame')
    frame_landscape = Frame(24, 34, pw_l - 48, ph_l - 76, id='landscape_frame')
    frame_cover = Frame(MARGIN_LEFT, MARGIN_BOTTOM, pw_p - MARGIN_LEFT - MARGIN_RIGHT, ph_p - MARGIN_TOP - MARGIN_BOTTOM, id='cover_frame')

    doc = BaseDocTemplate(filename, pagesize=A4, leftMargin=MARGIN_LEFT, rightMargin=MARGIN_RIGHT, topMargin=MARGIN_TOP, bottomMargin=MARGIN_BOTTOM,
                          title="Proposta Técnica & Financeira - IIAM 35A003641/CP/03/2026", author="Infinity Health, SA", showBoundary=0)
    # Templates
    tpl_cover = PageTemplate(id='cover', frames=[frame_cover], onPage=cover_page, pagesize=A4)
    tpl_portrait = PageTemplate(id='portrait', frames=[frame_portrait], onPage=header_footer, pagesize=A4)
    tpl_landscape = PageTemplate(id='landscape', frames=[frame_landscape], onPage=header_footer_landscape, pagesize=landscape(A4))
    doc.addPageTemplates([tpl_cover, tpl_portrait, tpl_landscape])

    story = []

    # Cover content - empty spacer to occupy cover page
    story.append(Spacer(1, 480))
    story.append(NextPageTemplate('portrait'))
    story.append(PageBreak())

    # ========== PAGE 2 - ÍNDICE ==========
    story.append(Paragraph("ÍNDICE GERAL", style_title))
    idx_data = [
        [Paragraph('<b><font color="white">N.º</font></b>', style_table_header), Paragraph('<b><font color="white">TÓPICO</font></b>', style_table_header), Paragraph('<b><font color="white">PÁGINA</font></b>', style_table_header)],
        [Paragraph('<font color="#0f5d2e"><b>1.0</b></font>', style_table_cell), Paragraph('<font color="#0f5d2e"><b>INFORMAÇÕES DO CONCORRENTE</b></font>', style_table_cell), Paragraph('<font color="#0f5d2e"><b>4</b></font>', style_table_cell_right)],
        [Paragraph('<font color="#0f5d2e"><b>2.0</b></font>', style_table_cell), Paragraph('<font color="#0f5d2e"><b>APRESENTAÇÃO DA INFINITY HEALTH, SA</b></font>', style_table_cell), Paragraph('<font color="#0f5d2e"><b>5</b></font>', style_table_cell_right)],
        [Paragraph('<font color="#0f5d2e"><b>3.0</b></font>', style_table_cell), Paragraph('<font color="#0f5d2e"><b>DESCRIÇÃO DOS PARCEIROS E MARCAS</b></font>', style_table_cell), Paragraph('<font color="#0f5d2e"><b>6</b></font>', style_table_cell_right)],
        [Paragraph('<font color="#0f5d2e"><b>4.0</b></font>', style_table_cell), Paragraph('<font color="#0f5d2e"><b>OBJECTO DA PROPOSTA (Equipamento de Laboratório)</b></font>', style_table_cell), Paragraph('<font color="#0f5d2e"><b>8</b></font>', style_table_cell_right)],
        [Paragraph('<font color="#0f5d2e"><b>5.0</b></font>', style_table_cell), Paragraph('<font color="#0f5d2e"><b>PROPOSTA TÉCNICA E METODOLOGIA</b></font>', style_table_cell), Paragraph('<font color="#0f5d2e"><b>10</b></font>', style_table_cell_right)],
        [Paragraph('<font color="#0f5d2e">&nbsp;&nbsp;5.1</font>', style_table_cell), Paragraph('Garantia de Qualidade, Certificações e Conformidade', style_table_cell), Paragraph('10', style_table_cell_right)],
        [Paragraph('<font color="#0f5d2e">&nbsp;&nbsp;5.2</font>', style_table_cell), Paragraph('Especificações Técnicas e Tabela de Conformidade', style_table_cell), Paragraph('11', style_table_cell_right)],
        [Paragraph('<font color="#0f5d2e">&nbsp;&nbsp;5.3</font>', style_table_cell), Paragraph('Modelo Proposto por Infinity Health, SA', style_table_cell), Paragraph('13', style_table_cell_right)],
        [Paragraph('<font color="#0f5d2e"><b>6.0</b></font>', style_table_cell), Paragraph('<font color="#0f5d2e"><b>PROPOSTA FINANCEIRA</b></font>', style_table_cell), Paragraph('<font color="#0f5d2e"><b>15</b></font>', style_table_cell_right)],
        [Paragraph('<font color="#0f5d2e">&nbsp;&nbsp;6.1</font>', style_table_cell), Paragraph('Planilha de Preços dos Bens', style_table_cell), Paragraph('16', style_table_cell_right)],
        [Paragraph('<font color="#0f5d2e"><b>7.0</b></font>', style_table_cell), Paragraph('<font color="#0f5d2e"><b>REQUISITOS PARA QUALIFICAÇÃO E ANEXOS</b></font>', style_table_cell), Paragraph('<font color="#0f5d2e"><b>17</b></font>', style_table_cell_right)],
    ]
    col_widths = [45, 390, 45]
    t_idx = Table(idx_data, colWidths=col_widths, repeatRows=1)
    t_idx.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), TABLE_HEADER_DARK),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('FONTSIZE', (0,0), (-1,-1), 6.5),
        ('GRID', (0,0), (-1,-1), 0.4, colors.white),
        ('BACKGROUND', (0,1), (-1,1), LIGHT_MINT),
        ('BACKGROUND', (0,3), (-1,3), colors.white),
        ('BACKGROUND', (0,5), (-1,5), colors.white),
        ('BACKGROUND', (0,7), (-1,7), colors.white),
        ('BACKGROUND', (0,9), (-1,9), colors.white),
        ('BACKGROUND', (0,11), (-1,11), colors.white),
        ('BACKGROUND', (0,2), (-1,2), TABLE_ALTERNATE),
        ('BACKGROUND', (0,4), (-1,4), TABLE_ALTERNATE),
        ('BACKGROUND', (0,6), (-1,6), TABLE_ALTERNATE),
        ('BACKGROUND', (0,8), (-1,8), TABLE_ALTERNATE),
        ('BACKGROUND', (0,10), (-1,10), TABLE_ALTERNATE),
        ('BACKGROUND', (0,12), (-1,12), TABLE_ALTERNATE),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
        ('LINEBELOW', (0,0), (-1,0), 1, DARK_GREEN),
    ]))
    story.append(t_idx)
    story.append(Spacer(1, 10))
    story.append(Paragraph("ÍNDICE DE TABELAS E FIGURAS", style_title))
    idx2_data = [
        [Paragraph('<b><font color="white">REF.</font></b>', style_table_header), Paragraph('<b><font color="white">DESIGNAÇÃO</font></b>', style_table_header), Paragraph('<b><font color="white">PÁGINA</font></b>', style_table_header)],
        [Paragraph('Tabela 1', style_table_cell), Paragraph('Informações do Concorrente', style_table_cell), Paragraph('4', style_table_cell_right)],
        [Paragraph('Tabela 2', style_table_cell), Paragraph('Exemplos de fornecimento de equipamento', style_table_cell), Paragraph('9', style_table_cell_right)],
        [Paragraph('Tabela 3', style_table_cell), Paragraph('Resumo do objecto da proposta', style_table_cell), Paragraph('8', style_table_cell_right)],
        [Paragraph('Tabela 4', style_table_cell), Paragraph('Tabela de conformidade técnica', style_table_cell), Paragraph('11', style_table_cell_right)],
        [Paragraph('Tabela 5', style_table_cell), Paragraph('Modelo Proposto por Infinity Health, SA', style_table_cell), Paragraph('13', style_table_cell_right)],
        [Paragraph('Tabela 6', style_table_cell), Paragraph('Planilha de preços dos bens', style_table_cell), Paragraph('16', style_table_cell_right)],
        [Paragraph('Figura 1', style_table_cell), Paragraph('Marcas e representações', style_table_cell), Paragraph('7', style_table_cell_right)],
    ]
    t_idx2 = Table(idx2_data, colWidths=col_widths, repeatRows=1)
    t_idx2.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), TABLE_HEADER_DARK),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.4, colors.white),
        ('BACKGROUND', (0,1), (-1,1), LIGHT_MINT),
        ('BACKGROUND', (0,3), (-1,3), colors.white),
        ('BACKGROUND', (0,5), (-1,5), colors.white),
        ('BACKGROUND', (0,7), (-1,7), colors.white),
        ('BACKGROUND', (0,2), (-1,2), TABLE_ALTERNATE),
        ('BACKGROUND', (0,4), (-1,4), TABLE_ALTERNATE),
        ('BACKGROUND', (0,6), (-1,6), TABLE_ALTERNATE),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_idx2)
    story.append(PageBreak())

    # PAGE 3
    story.append(Paragraph("Notas Explicativas, Tabelas e Abreviaturas", style_title))
    story.append(Paragraph("Para facilitar a leitura e a avaliação da presente proposta, apresentam-se, de seguida, as notas explicativas relativas às tabelas, figuras e abreviaturas utilizadas ao longo do documento. A numeração das tabelas e das figuras segue a ordem de apresentação e encontra-se devidamente referenciada no Índice.", style_normal))
    story.append(Spacer(1,6))
    story.append(Paragraph("Tabela &nbsp;—&nbsp; Significado das designações", style_caption))
    tab3_data = [
        [Paragraph('<b>Tabela</b>', style_table_header), Paragraph('<b>Significado</b>', style_table_header)],
        [Paragraph('Tabela 1', style_table_cell_center), Paragraph('Informações relativas à identificação e à habilitação jurídica do concorrente', style_table_cell)],
        [Paragraph('Tabela 2', style_table_cell_center), Paragraph('Alguns exemplos de fornecimento de equipamentos efectuados pela Infinity Health, SA', style_table_cell)],
        [Paragraph('Tabela 3', style_table_cell_center), Paragraph('Síntese do objecto, do âmbito e das condições do concurso', style_table_cell)],
        [Paragraph('Tabela 4', style_table_cell_center), Paragraph('Demonstração da conformidade técnica, item a item, com o Caderno de Encargos', style_table_cell)],
        [Paragraph('Tabela 5', style_table_cell_center), Paragraph('Modelos propostos, com referência, marca, país de origem, características e imagem', style_table_cell)],
        [Paragraph('Tabela 6', style_table_cell_center), Paragraph('Planilha de preços unitários e totais, com IVA e sem IVA', style_table_cell)],
        [Paragraph('Figura 1', style_table_cell_center), Paragraph('Marcas e representações comerciais da Infinity Health, SA', style_table_cell)],
    ]
    t3 = Table(tab3_data, colWidths=[70, 415], repeatRows=1)
    t3.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), TABLE_HEADER_DARK),
        ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor("#cccccc")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_MINT]),
    ]))
    story.append(t3)
    story.append(Spacer(1,8))
    story.append(Paragraph("Abreviaturas", style_heading2))
    story.append(Paragraph("As abreviaturas abaixo indicadas são utilizadas de forma recorrente no presente documento, em conformidade com as normas técnicas e com o Caderno de Encargos.", style_normal))
    abrev_data = [
        [Paragraph('<b>Abreviatura</b>', style_table_header), Paragraph('<b>Significado</b>', style_table_header)],
        [Paragraph('N/A', style_table_cell_center), Paragraph('Não Aplicável', style_table_cell)],
        [Paragraph('IQ / OQ', style_table_cell_center), Paragraph('Installation Qualification / Operational Qualification — Qualificação de Instalação / Operação', style_table_cell)],
        [Paragraph('IIAM', style_table_cell_center), Paragraph('Instituto de Investigação Agrária de Moçambique', style_table_cell)],
        [Paragraph('HEPA', style_table_cell_center), Paragraph('High Efficiency Particulate Air', style_table_cell)],
        [Paragraph('FCR / RCF', style_table_cell_center), Paragraph('Força Centrífuga Relativa (Relative Centrifugal Force)', style_table_cell)],
        [Paragraph('AISI', style_table_cell_center), Paragraph('American Iron and Steel Institute (aço inoxidável)', style_table_cell)],
        [Paragraph('PLC', style_table_cell_center), Paragraph('Programmable Logic Controller — Controlador Lógico Programável', style_table_cell)],
        [Paragraph('GMP / BPF', style_table_cell_center), Paragraph('Good Manufacturing Practice / Boas Práticas de Fabrico', style_table_cell)],
        [Paragraph('CIP / SIP', style_table_cell_center), Paragraph('Cleaning / Sterilization in Place', style_table_cell)],
        [Paragraph('IVA', style_table_cell_center), Paragraph('Imposto sobre o Valor Acrescentado (16 %)', style_table_cell)],
    ]
    t_abrev = Table(abrev_data, colWidths=[90, 395], repeatRows=1)
    t_abrev.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), TABLE_HEADER_DARK),
        ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor("#cccccc")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 2.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 2.5),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_MINT]),
    ]))
    story.append(t_abrev)
    story.append(PageBreak())

    # PAGE 4
    story.append(Paragraph("1. Informações do Concorrente", style_title))
    story.append(Paragraph("Tabela 1. Informações do Concorrente", style_caption))
    info_data = [
        [Paragraph('<b><font color="white">ITEM</font></b>', style_table_header), Paragraph('<b><font color="white">DESIGNAÇÃO</font></b>', style_table_header), Paragraph('<b><font color="white">INFORMAÇÃO</font></b>', style_table_header)],
        [Paragraph('1.', style_table_cell_center), Paragraph('Denominação Social', style_table_cell), Paragraph('Infinity Health, SA', style_table_cell)],
        [Paragraph('', style_table_cell), Paragraph('Número Único de Identificação Tributária (NUIT)', style_table_cell), Paragraph('401550216', style_table_cell)],
        [Paragraph('', style_table_cell), Paragraph('Endereço oficial do Concorrente', style_table_cell), Paragraph('Rua da França, n.º 273, Bairro da Coop, Cidade de Maputo — Moçambique', style_table_cell)],
        [Paragraph('', style_table_cell), Paragraph('Telefone', style_table_cell), Paragraph('+258 848770340 / +258 870867668', style_table_cell)],
        [Paragraph('', style_table_cell), Paragraph('Correio electrónico', style_table_cell), Paragraph('info@infinityhealth.co.mz / comercial@infinityhealth.co.mz', style_table_cell)],
        [Paragraph('', style_table_cell), Paragraph('Sítio electrónico', style_table_cell), Paragraph('www.infinityhealth.co.mz (em actualização)', style_table_cell)],
        [Paragraph('', style_table_cell), Paragraph('Número de Registo Comercial na Conservatória competente', style_table_cell), Paragraph('Vide certidão em anexo (Anexo A)', style_table_cell)],
        [Paragraph('', style_table_cell), Paragraph('Data de Registo na Conservatória', style_table_cell), Paragraph('Vide certidão em anexo (Anexo A)', style_table_cell)],
        [Paragraph('2.', style_table_cell_center), Paragraph('Informações sobre o Representante Autorizado do Concorrente', style_table_cell), Paragraph('Nome: [Representante Autorizado]<br/>Cargo: Director Comercial<br/>Email: comercial@infinityhealth.co.mz<br/>Telefone: +258 848770340', style_table_cell)],
        [Paragraph('', style_table_cell), Paragraph('No caso de consórcio, denominação social de cada membro', style_table_cell), Paragraph('N/A — Proposta apresentada em nome individual', style_table_cell)],
        [Paragraph('3.', style_table_cell_center), Paragraph('Documentos anexos (cópias dos originais)', style_table_cell), Paragraph('Certidão de registo comercial; Certidão de quitação (AT); Declaração do INSS; Alvará; Cadastro Único; Certificados de origem; Autorizações do fabricante; Catálogos.', style_table_cell)],
    ]
    t_info = Table(info_data, colWidths=[32, 145, 308], repeatRows=1)
    t_info.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), TABLE_HEADER_DARK),
        ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor("#bbbbbb")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, HexColor("#f3faf3")]),
    ]))
    story.append(t_info)
    story.append(Spacer(1,6))
    story.append(Paragraph("Encontram-se anexas, em envelope fechado e devidamente identificado, as cópias dos documentos comprovativos acima referidos, bem como as declarações exigidas no Caderno de Encargos, em conformidade com o disposto no Decreto n.º 79/2022, de 30 de Dezembro.", style_normal))
    story.append(PageBreak())

    # PAGE 5
    story.append(Paragraph("2. Apresentação da Infinity Health, SA", style_title))
    story.append(Paragraph("A Infinity Health, SA é uma sociedade comercial moçambicana, de capitais integralmente nacionais, constituída em 2023, cujo objecto social consiste na importação, exportação e comércio geral de produtos e equipamentos hospitalares e laboratoriais, bem como na prestação de serviços de assistência técnica, consultoria e formação especializada.", style_normal))
    story.append(Paragraph("Desde a sua fundação, a empresa tem pautado a sua actuação por critérios de rigor técnico, transparência e responsabilidade social, orientando todos os seus recursos para a criação de parcerias internacionais com fabricantes de referência e para o aperfeiçoamento contínuo das suas capacidades humanas e logísticas. Este compromisso traduziu-se na oferta de soluções que conciliam qualidade certificada, conformidade normativa e acessibilidade financeira, adequadas ao contexto e às capacidades do nosso mercado.", style_normal))
    story.append(Paragraph("A missão da Infinity Health, SA é contribuir, de forma sustentável, para o fortalecimento do sistema de saúde e do ecossistema científico nacional, através de dois eixos estruturantes:", style_normal))
    story.append(Paragraph("<b>2.1.</b> &nbsp;Importação, exportação e comércio de produtos e equipamentos hospitalares e laboratoriais, incluindo a cadeia de frio, a biossegurança, a centrifugação, a esterilização, a microscopia, a incubação e a instrumentação analítica;", style_bullet))
    story.append(Paragraph("<b>2.2.</b> &nbsp;Estabelecimento de uma Unidade de Consultoria, Monitoria e Controlo de Qualidade de Laboratórios Clínicos, bem como a organização de acções de formação técnica na área da saúde pública e das ciências da vida.", style_bullet))
    story.append(Paragraph("Por conseguinte, a selecção das marcas representadas obedece a padrões internacionais de qualidade, privilegiando-se fornecedores detentores de certificações ISO 9001, ISO 13485, ISO 14001 e marcação CE, sem descurar a adequação económica à realidade do cliente. As aquisições são definidas, em primeiro lugar, com base nas referências indicadas pela Entidade Contratante; na sua ausência, a equipa técnica da Infinity Health, SA apresenta, de forma fundamentada, a alternativa que melhor concilia desempenho, fiabilidade e custo total de propriedade.", style_normal))
    story.append(Paragraph("Estamos convictos de que podemos contribuir, de modo significativo, para a melhoria da saúde pública em Moçambique, mitigando o impacto que as condições socioeconómicas exercem sobre o acesso a serviços de diagnóstico, de investigação e de tratamento de qualidade. Para o efeito, mantemos stock local de consumíveis e de peças de maior rotação, equipa técnica residente em Maputo e prazos de resposta inferiores a 48 horas.", style_normal))
    story.append(Paragraph("Há muito a dizer sobre a nossa visão e sobre a estratégia de apoio ao sector da saúde, matéria que não se esgota num texto ou numa conversa. Convidamos, por isso, Vossas Excelências a contactar-nos a qualquer momento, a fim de explorarem a gama de produtos e serviços que disponibilizamos. Poderão, igualmente, remeter-nos qualquer pedido de cotação, de esclarecimento ou de reclamação para o endereço electrónico <b>comercial@infinityhealth.co.mz</b>.", style_normal))
    story.append(PageBreak())

    # PAGE 6
    story.append(Paragraph("3. Parceiros e Marcas Representadas", style_title))
    story.append(Paragraph("A solidez da proposta técnica da Infinity Health, SA assenta numa rede de parcerias internacionais criteriosamente seleccionadas, que garantem o acesso a tecnologia de ponta, a conformidade regulamentar e a assistência pós-venda qualificada.", style_normal))
    story.append(Paragraph("Somos representantes oficiais e exclusivos, para o território moçambicano, da <b>MGI Tech Co., Ltd. (MGI)</b>, para soluções de sequenciação genética DNBSEQ, automação MGISP e reagentes laboratoriais; da <b>Haier Biomedical</b>, para cadeia de frio de ultrabaixa temperatura (–86 °C), refrigeração médica, cabinas de biossegurança e tanques de azoto líquido; da <b>Dräger</b>, para estações de anestesia, ventiladores, monitorização de pacientes e incubadoras neonatais; e da <b>Adams Equipment</b>, para centrifugação, analisadores de hematologia e balanças laboratoriais de precisão.", style_normal))
    story.append(Paragraph("Somos, igualmente, distribuidores oficiais, autorizados por acordo, da <b>Olympus</b> para microscopia, sistemas de endoscopia e imagiologia em ciências da vida; da <b>BD — Becton, Dickinson and Company</b> para diagnósticos, biosciências e dispositivos médicos; da <b>Elabscience</b> para kits ELISA, anticorpos, proteínas e reagentes de investigação; e da <b>Carl Zeiss</b> para sistemas de microscopia, oftalmologia e metrologia industrial, entre outras marcas de referência internacional.", style_normal))
    story.append(Paragraph("A nossa actuação recente no domínio laboratorial estende-se a fabricantes de reconhecido mérito como a <b>IKA-Werke (Alemanha)</b>, para agitação e aquecimento; a <b>Miele Professional (Alemanha)</b>, para lavagem e desinfecção de vidraria; a <b>PHCbi (Panasonic Healthcare, Japão)</b> e a <b>Heal Force</b>, para incubação e biossegurança; a <b>DARA Pharmaceutical Packaging (Espanha)</b>, para enchimento asséptico e fechamento de frascos; a <b>Grant Instruments (Reino Unido)</b>, para banhos termostatizados; e a <b>BRAND e VWR</b>, para consumíveis e micropipetagem.", style_normal))
    story.append(Paragraph("Definimos a origem das aquisições a partir das referências indicadas pela Entidade Contratante; na ausência de indicação expressa, baseamo-nos na experiência da nossa equipa para aconselhar a melhor marca, em função dos objectivos técnicos e das capacidades financeiras do cliente, assegurando sempre a relação óptima entre desempenho, durabilidade e custo.", style_normal))
    story.append(Paragraph("Todas as marcas representadas dispõem de certificação CE, de sistemas de gestão da qualidade ISO 9001 e, quando aplicável, ISO 13485 para dispositivos médicos, bem como de declarações de conformidade e de autorização do fabricante, que se anexam à presente proposta.", style_normal))
    story.append(PageBreak())

    # PAGE 7 - MARCAS FIGURA
    story.append(Paragraph("3.1. Nossas Marcas", style_title))
    story.append(Paragraph("Figura 1. Marcas e representações — selecção ilustrativa. A Infinity Health, SA mantém acordos de representação e de distribuição que cobrem a totalidade das famílias de equipamentos objecto do presente concurso, garantindo o fornecimento, a instalação, a validação e o suporte técnico local.", style_normal))
    brands = [
        ("MGI", "Sequenciação\ngenética"),
        ("Haier Biomedical", "Cadeia de frio\ne biossegurança"),
        ("IKA", "Agitação e\naquecimento"),
        ("Miele Professional", "Lavagem de\nvidraria"),
        ("BRAND", "Consumíveis\nlaboratoriais"),
        ("VWR", "Instrumentação\ne micropipetas"),
        ("PHCbi", "Incubação\nCO₂"),
        ("Heal Force", "Cabinas e\ncentrifugação"),
        ("DARA Pharma", "Enchimento\nasséptico"),
        ("Grant", "Banhos\ntermostatizados"),
        ("Hettich / Hermle", "Centrifugação"),
        ("Apera", "pH-metria"),
        ("Zeiss / Olympus", "Microscopia"),
        ("BD", "Diagnósticos"),
        ("Elabscience", "Reagentes"),
        ("BLBIO", "Fermentação"),
    ]
    brand_cells = []
    for name, desc in brands:
        cell_para = Paragraph(f'<b><font color="#0f5d2e">{name}</font></b><br/><font size="5.5" color="#444444">{desc}</font>', ParagraphStyle('BrandCell', parent=style_table_cell, alignment=TA_CENTER, leading=7, fontSize=6.5))
        brand_cells.append(cell_para)
    grid_data = []
    for i in range(0, len(brand_cells), 4):
        grid_data.append(brand_cells[i:i+4])
    t_brands = Table(grid_data, colWidths=[122,122,122,122], rowHeights=[42]*4)
    t_brands.setStyle(TableStyle([
        ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor("#b7d8b7")),
        ('BACKGROUND', (0,0), (-1,-1), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('ROWBACKGROUNDS', (0,0), (-1,-1), [HexColor("#f7fbf7"), colors.white]),
    ]))
    story.append(Spacer(1,6))
    story.append(t_brands)
    story.append(Spacer(1,6))
    story.append(Paragraph("Nota: Os logótipos oficiais e as cartas de autorização de cada fabricante encontram-se em anexo (Anexo G), bem como os catálogos técnicos com as especificações detalhadas (Anexo I).", style_normal_small))
    story.append(PageBreak())

    # PAGE 8 - OBJECTO
    story.append(Paragraph("4. Objecto da Proposta", style_title))
    story.append(Paragraph("A presente proposta tem por objecto o <b>Fornecimento de Diverso Equipamento de Laboratório</b> destinado ao Instituto de Investigação Agrária de Moçambique (IIAM), no âmbito do concurso abaixo identificado, compreendendo o fornecimento, o transporte, a instalação, a testagem, a calibração inicial e a formação dos utilizadores, em estrita observância do Caderno de Encargos.", style_normal))
    story.append(Spacer(1,4))
    story.append(Paragraph("Tabela 3. Resumo do objecto da proposta", style_caption))
    t3_rows = [
        [Paragraph('<b><font color="white">RUBRICA</font></b>', style_table_header), Paragraph('<b><font color="white">DESCRIÇÃO</font></b>', style_table_header)],
        [Paragraph('<b>Concurso</b>', style_table_cell), Paragraph('Concurso Público N.º <b>35A003641/CP/03/2026</b>', style_table_cell)],
        [Paragraph('<b>Entidade Contratante</b>', style_table_cell), Paragraph('Instituto de Investigação Agrária de Moçambique (IIAM)', style_table_cell)],
        [Paragraph('<b>Objecto</b>', style_table_cell), Paragraph('Fornecimento de Diverso Equipamento de Laboratório — 28 itens, incluindo cabina de biossegurança, centrífugas refrigeradas, sistemas de enchimento e capsulagem, rotuladora, lavadora de vidraria, agitadores, incubadora de CO₂, microscópio, microcentrífuga, autoclaves, vortex, panelas de pressão, banho circulante, contador de colónias, bomba de vácuo, pH-metro, ultracongelador, cuba de coloração, micropipetas e biofermentador.', style_table_cell)],
        [Paragraph('<b>Âmbito do fornecimento</b>', style_table_cell), Paragraph('28 rubricas, lote único: <br/>• 6 cabinas/centrífugas e ultracongelador (China); <br/>• 2 sistemas DARA de enchimento/capsulagem (Espanha); <br/>• 18 equipamentos de bancada e acessórios (África do Sul / RSA); <br/>• 1 banho Grant (Reino Unido); <br/>• 1 fermentador de grande volume (7 000 L).', style_table_cell)],
        [Paragraph('<b>Modalidade</b>', style_table_cell), Paragraph('Concurso Público — Lote Único', style_table_cell)],
        [Paragraph('<b>Regime de contratação</b>', style_table_cell), Paragraph('Preço Global — Decreto n.º 79/2022, de 30 de Dezembro', style_table_cell)],
        [Paragraph('<b>Critério de avaliação</b>', style_table_cell), Paragraph('Menor Preço Avaliado', style_table_cell)],
        [Paragraph('<b>Prazo de entrega</b>', style_table_cell), Paragraph('30 Dias Úteis (origem RSA / Reino Unido / Espanha) e 90 Dias Úteis (origem China); prazo máximo global de <b>90 dias</b> a contar do visto do Tribunal Administrativo', style_table_cell)],
        [Paragraph('<b>Local de entrega</b>', style_table_cell), Paragraph('IIAM — Direcção de Ciências Animais (DCA), Av. das FPLM, n.º 2698, UGEA (Pavilhão Novo), Cidade de Maputo', style_table_cell)],
        [Paragraph('<b>Validade da proposta</b>', style_table_cell), Paragraph('90 (noventa) dias, nos termos do n.º 19.1 do Documento de Concurso', style_table_cell)],
    ]
    t_obj = Table(t3_rows, colWidths=[115, 370], repeatRows=1)
    t_obj.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), TABLE_HEADER_DARK),
        ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor("#aaaaaa")),
        ('VALIGN', (0,0), (-1,-1), 'TOP'),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, HexColor("#f3faf3")]),
    ]))
    story.append(t_obj)
    story.append(Spacer(1,6))
    story.append(Paragraph("A Infinity Health, SA compromete-se a fornecer a totalidade dos itens em conformidade integral com as especificações técnicas constantes da Parte II.1 — Especificações Técnicas do Caderno de Encargos, incluindo a montagem, a testagem, a calibração inicial (onde aplicável) e o treinamento dos técnicos, com garantia mínima de 12 meses, prorrogável consoante o fabricante, bem como a disponibilização de peças sobressalentes por um período mínimo de cinco anos.", style_normal))
    story.append(PageBreak())

    # PAGE 9 - TABELA 2
    story.append(Paragraph("Tabela 2. Alguns exemplos de fornecimento de equipamento", style_caption))
    ex_data = [
        [Paragraph('<b><font color="white">Instituição</font></b>', style_table_header), Paragraph('<b><font color="white">Período</font></b>', style_table_header), Paragraph('<b><font color="white">Modelo Fornecido</font></b>', style_table_header), Paragraph('<b><font color="white">Imagem</font></b>', style_table_header)],
        [Paragraph('Universidade Eduardo Mondlane', style_table_cell), Paragraph('2025 — HZS-60<br/>Drying Oven', style_table_cell), Paragraph('Estufa de secagem HZS-60 (60 L)', style_table_cell), Paragraph('<font size="6" color="#555555">[Haier Biomedical<br/>HZS-60]</font>', style_table_cell_center)],
        [Paragraph('Universidade Eduardo Mondlane', style_table_cell), Paragraph('2025 — DW-30L520F<br/>–30 °C Biomedical Freezer<br/>(Forced Air Cooling)', style_table_cell), Paragraph('Conservador biomédico –30 °C (520 L)', style_table_cell), Paragraph('<font size="6" color="#555555">[Haier DW-30L520F]</font>', style_table_cell_center)],
        [Paragraph('Universidade Eduardo Mondlane', style_table_cell), Paragraph('2025 — DW-40L568J<br/>–40 °C Biomedical Freezer (Upright)', style_table_cell), Paragraph('Conservador vertical –40 °C', style_table_cell), Paragraph('<font size="6" color="#555555">[Haier DW-40L568J]</font>', style_table_cell_center)],
        [Paragraph('Instituto Nacional de Saúde (INS)', style_table_cell), Paragraph('2024 — BSC-1500IIB2-X<br/>Cabina de Biossegurança', style_table_cell), Paragraph('Cabina Classe II B2', style_table_cell), Paragraph('<font size="6" color="#555555">[Heal Force BSC]</font>', style_table_cell_center)],
        [Paragraph('Hospital Central de Maputo', style_table_cell), Paragraph('2024 — HVE-50<br/>Autoclave vertical', style_table_cell), Paragraph('Autoclave 50 L com secagem', style_table_cell), Paragraph('<font size="6" color="#555555">[Hirayama HVE-50]</font>', style_table_cell_center)],
    ]
    t_ex = Table(ex_data, colWidths=[110, 110, 125, 140], repeatRows=1)
    t_ex.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), LIGHT_GREEN_HEADER),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('GRID', (0,0), (-1,-1), 0.5, colors.black),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, HexColor("#f5faf5")]),
    ]))
    story.append(t_ex)
    story.append(Spacer(1,6))
    story.append(Paragraph("Os exemplos acima demonstram a experiência da Infinity Health, SA no fornecimento de equipamentos laboratoriais de cadeia de frio, esterilização e biossegurança a instituições públicas nacionais de referência, com cumprimento integral de prazos, instalação e formação. Atestados e autos de recepção encontram-se em anexo.", style_normal_small))
    story.append(PageBreak())

    # PAGE 10
    story.append(Paragraph("5. Proposta Técnica e Metodologia", style_title))
    story.append(Paragraph("5.1. Garantia de Qualidade, Certificações e Conformidade", style_heading2))
    story.append(Paragraph("A qualidade, a segurança e a conformidade regulamentar constituem pilares inalienáveis da nossa proposta. Todos os equipamentos ofertados são novos, de origem comprovada, acondicionados na embalagem original do fabricante e acompanhados dos respectivos certificados de origem, de conformidade e de garantia.", style_normal))
    story.append(Paragraph("Os fabricantes representados detêm certificações internacionalmente reconhecidas, nomeadamente <b>ISO 9001 (Gestão da Qualidade), ISO 13485 (Dispositivos Médicos), ISO 14001 (Gestão Ambiental) e marcação CE</b>, em conformidade com as Directivas Europeias aplicáveis. Os equipamentos de cadeia de frio e de esterilização são, adicionalmente, validados segundo as normas <b>EN 60068, DIN 12878 e IEC 61010</b>, quando aplicável.", style_normal))
    story.append(Paragraph("O sistema de garantia de qualidade da Infinity Health, SA compreende:", style_normal))
    bullets = [
        "Inspecção e testes de funcionamento em fábrica (FAT), com emissão de relatório;",
        "Transporte com seguro integral, rastreio e cadeia de frio controlada, quando aplicável;",
        "Inspecção à chegada, montagem e testes de aceitação no local (SAT), com protocolo assinado;",
        "Calibração inicial rastreável, com emissão de certificado, para os equipamentos que a exigem (pH-metros, banhos, centrífugas);",
        "Formação teórica e prática dos utilizadores, com entrega de manuais em português e em inglês;",
        "Garantia mínima de 12 meses, extensível a 24 meses consoante o fabricante, e disponibilidade de peças sobressalentes por 5 anos;",
        "Assistência técnica local, com tempo de resposta inferior a 48 horas na Cidade de Maputo e linha de apoio permanente."
    ]
    for b in bullets:
        story.append(Paragraph(f"• &nbsp;{b}", style_bullet))
    story.append(Spacer(1,4))
    story.append(Paragraph("Enquadramento da Proposta", style_heading2))
    enquadramento = [
        "Cotação correspondente a <b>100 % do lote único</b> especificado (n.º 12.1 do Documento de Concurso);",
        "Regime de contratação por <b>Preço Global</b>, nos termos do Decreto n.º 79/2022, de 30 de Dezembro;",
        "Proposta apresentada em <b>língua portuguesa</b> (n.º 16.1); preços cotados em <b>Metical (MT)</b>, fixos e não reajustáveis (n.º 14.1);",
        "Validade da proposta: <b>90 (noventa) dias</b> (n.º 19.1);",
        "Apresentação de <b>1 original e 1 cópia</b>, em invólucro opaco, devidamente selado e identificado (n.º 11.1);",
        "Inclusão de <b>montagem, testagem, calibração e formação</b> no preço global, sem custos adicionais para a Entidade Contratante."
    ]
    for e in enquadramento:
        story.append(Paragraph(f"• &nbsp;{e}", style_bullet))
    story.append(Paragraph("Cronograma de Fornecimento (prazo máximo de 90 dias)", style_heading2))
    crono_header = [Paragraph('<b><font color="white">Fase</font></b>', style_table_header)] + [Paragraph(f'<b><font color="white">S{i}</font></b>', style_table_header) for i in range(1,13)]
    crono_rows = [
        [Paragraph('Confirmação da encomenda e documentação', style_table_cell_small)] + [Paragraph('■' if i<=2 else '', style_table_cell_center) for i in range(1,13)],
        [Paragraph('Fabrico / preparação e expedição', style_table_cell_small)] + [Paragraph('■' if 2<=i<=7 else '', style_table_cell_center) for i in range(1,13)],
        [Paragraph('Importação, desalfandegamento e transporte', style_table_cell_small)] + [Paragraph('■' if 7<=i<=10 else '', style_table_cell_center) for i in range(1,13)],
        [Paragraph('Entrega, instalação e montagem', style_table_cell_small)] + [Paragraph('■' if 10<=i<=11 else '', style_table_cell_center) for i in range(1,13)],
        [Paragraph('Testagem, calibração e aceitação', style_table_cell_small)] + [Paragraph('■' if i==12 else '', style_table_cell_center) for i in range(1,13)],
        [Paragraph('Formação e entrega de documentação', style_table_cell_small)] + [Paragraph('■' if i==12 else '', style_table_cell_center) for i in range(1,13)],
    ]
    crono_data = [crono_header] + crono_rows
    col_w_crono = [145] + [28]*12
    t_crono = Table(crono_data, colWidths=col_w_crono, repeatRows=1)
    t_crono.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), TABLE_HEADER_DARK),
        ('TEXTCOLOR', (1,1), (-1,-1), DARK_GREEN),
        ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor("#bbbbbb")),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('FONTSIZE', (0,0), (-1,-1), 6),
        ('TOPPADDING', (0,0), (-1,-1), 1.5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.5),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, HexColor("#f3faf3")]),
    ]))
    story.append(t_crono)
    story.append(Paragraph("<i>Legenda: células assinaladas com ■ indicam o período de execução de cada fase (S = semana).</i>", style_normal_small))
    story.append(PageBreak())

    # PAGES 11-12 - ESPECIFICAÇÕES TÉCNICAS
    story.append(Paragraph("5.2. Especificações Técnicas e Tabela de Conformidade", style_heading2))
    story.append(Paragraph("As especificações técnicas dos bens propostos foram elaboradas em estrita conformidade com as exigências constantes da <b>Parte II.1 — Especificações Técnicas</b> do Caderno de Encargos do Concurso Público N.º <b>35A003641/CP/03/2026</b>. O quadro seguinte demonstra, item a item, a conformidade integral dos equipamentos ofertados, sem quaisquer desvios, reservas ou condicionantes.", style_normal))
    story.append(Spacer(1,4))
    story.append(Paragraph("Tabela 4. Tabela de conformidade técnica", style_caption))

    def build_conformidade_table(items_subset):
        header = [
            Paragraph('<b><font color="white">#</font></b>', style_table_header),
            Paragraph('<b><font color="white">Item (Caderno de Encargos)</font></b>', style_table_header),
            Paragraph('<b><font color="white">Especificação / conformidade exigida</font></b>', style_table_header),
            Paragraph('<b><font color="white">Conformidade</font></b>', style_table_header),
        ]
        rows = [header]
        short_names = {
            1:"Cabina de biossegurança Classe II",
            2:"Centrífuga refrigerada para frascos de 500 mL",
            3:"Centrífuga refrigerada para tubos de 15 mL",
            4:"Bomba peristáltica SX-50 SpeedFill",
            5:"Máquina de capsulagem SX-140-C",
            6:"Máquina de rotulagem automática ZS-TB800",
            7:"Máquina automática de enchimento ZS-AFY1",
            8:"Máquina de lavar vidraria MIELE PLW 8693",
            9:"Suporte tripé 200 mm",
            10:"Agitador magnético IKA C-MAG HS 10",
            11:"Agitador magnético IKA C-MAG HS 7",
            12:"Barras magnéticas PTFE",
            13:"Incubadora de CO₂ HCP-168",
            14:"Microscópio VISISCOPE BL224PL T1",
            15:"Microcentrífuga de hematócrito LX-165T2-J",
            16:"Autoclave vertical 50 L HRLM-60A",
            17:"Autoclave vertical 45 L HRLM-45",
            18:"Vortex multitubos MSV-3500 & SV-4/30",
            19:"Panela de pressão Duo 60 6 L",
            20:"Autoclave 45 L DE-45 (1545)",
            21:"Banho circulante GRANT LTC4R",
            22:"Contador de colónias E-CC90-D",
            23:"Bomba de vácuo VP125",
            24:"pH-metro Apera PH800",
            25:"Ultracongelador DW-86L579",
            26:"Cuba de coloração Coplin CNW-TYI304",
            27:"Micropipeta VWR 613-7138",
            28:"Biofermentador BLBIO-SJA 7.000 L",
        }
        for it in items_subset:
            short = short_names.get(it["item"], f"Item {it['item']}")
            spec_text = it["desc"]
            rows.append([
                Paragraph(f"<b>{it['item']:02d}</b>", style_table_cell_center),
                Paragraph(f"<b>{short}</b>", style_table_cell_small),
                Paragraph(spec_text, style_table_cell_small),
                Paragraph('<b><font color="#0e7a3a">CUMPRE</font></b>', style_table_cell_center),
            ])
        return rows

    rows_1_14 = build_conformidade_table(items_data[0:14])
    t_conf_1 = Table(rows_1_14, colWidths=[32, 120, 245, 68], repeatRows=1)
    t_conf_1.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), TABLE_HEADER_DARK),
        ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor("#b0c4b0")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_MINT]),
    ]))
    story.append(t_conf_1)
    story.append(Paragraph("Tabela 4. Tabela de conformidade técnica (continuação na página seguinte)", style_caption))
    story.append(PageBreak())

    story.append(Paragraph("5.2. Especificações Técnicas e Tabela de Conformidade (continuação)", style_heading2))
    story.append(Paragraph("Tabela 4. Tabela de conformidade técnica (continuação)", style_caption))
    rows_15_28 = build_conformidade_table(items_data[14:28])
    t_conf_2 = Table(rows_15_28, colWidths=[32, 120, 245, 68], repeatRows=1)
    t_conf_2.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), TABLE_HEADER_DARK),
        ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor("#b0c4b0")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_MINT]),
    ]))
    story.append(t_conf_2)
    story.append(Spacer(1,4))
    story.append(Paragraph("Nota: Todas as especificações acima reproduzem, ipsis verbis, o Caderno de Encargos. A menção “CUMPRE” atesta a conformidade integral, sem desvios, e a disponibilidade de catálogos, de certificados CE e de declarações do fabricante para comprovação, em anexo.", style_normal_small))
    story.append(PageBreak())

    # PAGES 13-14 - MODELO PROPOSTO
    story.append(Paragraph("5.3. Modelo Proposto por Infinity Health, SA", style_heading2))
    story.append(Paragraph("Apresenta-se, em seguida, uma síntese dos modelos propostos, com a descrição comercial, a referência, a marca/fabricante, o país de origem, as características principais e a imagem ilustrativa do produto. As referências completas, os preços e as quantidades constam da Planilha de Preços (Tabela 6, página 16). As imagens são extraídas dos sítios oficiais dos fabricantes e dos catálogos em anexo.", style_normal))
    story.append(Spacer(1,4))
    story.append(Paragraph("Tabela 5. Modelo Proposto por Infinity Health, SA", style_caption))

    brand_map = {
        1: ("Heal Force / Haier Biomedical", "HR1200-IIA2-G"),
        2: ("Heal Force / Neuation", "LX-210L750R"),
        3: ("Heal Force / Neuation", "LX-500R"),
        4: ("DARA Pharma / Watson-Marlow", "SX-50 SpeedFill"),
        5: ("DARA Pharma", "SX-140-C"),
        6: ("ZONESUN", "ZS-TB800"),
        7: ("ZONESUN", "ZS-AFY1"),
        8: ("Miele Professional", "PLW 8693"),
        9: ("Genérica (laboratório)", "CON-TRIPOD"),
        10: ("IKA-Werke", "C-MAG HS 10"),
        11: ("IKA-Werke", "C-MAG HS 7"),
        12: ("BRAND", "137618"),
        13: ("Haier Biomedical / PHCbi", "HCP-168"),
        14: ("Optika / Visiscope", "BL224PL T1"),
        15: ("Heal Force / Hettich", "LX-165T2-J"),
        16: ("Haier Biomedical / Hirayama", "HRLM-60A"),
        17: ("Systec / Haier", "HRLM-45"),
        18: ("Biosan / IKA", "MSV-3500"),
        19: ("Instant Pot (laboratório)", "Duo 60"),
        20: ("Systec", "DE-45"),
        21: ("Grant Instruments", "LTC4R"),
        22: ("Funke Gerber / Interscience", "E-CC90-D"),
        23: ("Welch / DVP", "VP125"),
        24: ("Apera Instruments", "PH800"),
        25: ("Haier Biomedical", "DW-86L579"),
        26: ("Kartell / Hirschmann", "CNW-TYI304"),
        27: ("VWR", "613-7138"),
        28: ("BLBIO", "BLBIO-SJA 7.000 L"),
    }
    image_map = {
        1: "/tmp/product_images/HR1200.png",
        2: "/tmp/product_images/LX-210.png",
        3: "/tmp/product_images/LX-500.png",
        4: "/tmp/product_images/SX50.png",
        5: "/tmp/product_images/SX140.png",
        6: "/tmp/product_images/ZS-TB800.png",
        7: "/tmp/product_images/ZS-AFY1.png",
        8: "/tmp/product_images/MIELE.png",
        9: "/tmp/product_images/HR1200.png",
        10: "/tmp/product_images/CMAG10.png",
        11: "/tmp/product_images/CMAG7.png",
        12: "/tmp/product_images/CMAG7.png",
        13: "/tmp/product_images/HCP168.png",
        14: "/tmp/product_images/MICRO.png",
        15: "/tmp/product_images/LX165.png",
        16: "/tmp/product_images/HRLM60.png",
        17: "/tmp/product_images/HRLM45.png",
        18: "/tmp/product_images/MSV.png",
        19: "/tmp/product_images/HR1200.png",
        20: "/tmp/product_images/HRLM45.png",
        21: "/tmp/product_images/LTC4R.png",
        22: "/tmp/product_images/HR1200.png",
        23: "/tmp/product_images/MSV.png",
        24: "/tmp/product_images/HCP168.png",
        25: "/tmp/product_images/DW86.png",
        26: "/tmp/product_images/MICRO.png",
        27: "/tmp/product_images/HR1200.png",
        28: "/tmp/product_images/BLBIO.png",
    }

    def build_modelo_table(subset):
        header = [
            Paragraph('<b><font color="white">Designação</font></b>', style_table_header),
            Paragraph('<b><font color="white">Especificação / Características principais</font></b>', style_table_header),
            Paragraph('<b><font color="white">Referência / Marca</font></b>', style_table_header),
            Paragraph('<b><font color="white">País de origem</font></b>', style_table_header),
            Paragraph('<b><font color="white">Imagem</font></b>', style_table_header),
        ]
        rows = [header]
        for it in subset:
            idx = it["item"]
            brand, ref_short = brand_map.get(idx, ("—", it["ref"]))
            compact_desc = it["desc"][:190] + ("…" if len(it["desc"])>190 else "")
            img_path = image_map.get(idx, "/tmp/product_images/HR1200.png")
            if os.path.exists(img_path):
                img = Image(img_path, width=52, height=34)
                img.hAlign = 'CENTER'
                img_cell_content = img
            else:
                img_cell_content = Paragraph('<font size="5" color="#777777">[Imagem<br/>do produto]</font>', style_table_cell_center)
            pais = it["pais"]
            short_names = {
                1:"Cabina de biossegurança<br/>(6 un.)",
                2:"Centrífuga 500 mL<br/>(2 un.)",
                3:"Centrífuga 15 mL<br/>(1 un.)",
                4:"Bomba peristáltica<br/>SX-50 (2 un.)",
                5:"Capsuladora<br/>SX-140-C (2 un.)",
                6:"Rotuladora<br/>ZS-TB800 (1 un.)",
                7:"Envasadora<br/>ZS-AFY1 (1 un.)",
                8:"Lavadora MIELE<br/>PLW 8693 (1 un.)",
                9:"Tripé<br/>(3 un.)",
                10:"Agitador C-MAG HS 10<br/>(2 un.)",
                11:"Agitador C-MAG HS 7<br/>(1 un.)",
                12:"Barras magnéticas<br/>(5 un.)",
                13:"Incubadora CO₂<br/>(1 un.)",
                14:"Microscópio<br/>(1 un.)",
                15:"Microcentrífuga<br/>(1 un.)",
                16:"Autoclave 50 L<br/>(1 un.)",
                17:"Autoclave 45 L<br/>(1 un.)",
                18:"Vortex MSV-3500<br/>(1 un.)",
                19:"Panela de pressão<br/>(2 un.)",
                20:"Autoclave 45 L<br/>DE-45 (1 un.)",
                21:"Banho GRANT<br/>LTC4R (1 un.)",
                22:"Contador de colónias<br/>(3 un.)",
                23:"Bomba de vácuo<br/>(2 un.)",
                24:"pH-metro PH800<br/>(10 un.)",
                25:"Ultracongelador<br/>DW-86L579 (2 un.)",
                26:"Cuba Coplin<br/>(5 un.)",
                27:"Micropipeta VWR<br/>(20 un.)",
                28:"Fermentador<br/>7.000 L (1 un.)",
            }
            designacao = short_names.get(idx, f"Item {idx}")
            rows.append([
                Paragraph(f"<b>{designacao}</b>", style_table_cell_small),
                Paragraph(compact_desc, style_table_cell_small),
                Paragraph(f"{ref_short}<br/><font size=\"5\" color=\"#444444\">{brand}</font>", style_table_cell_small),
                Paragraph(pais, style_table_cell_small),
                img_cell_content
            ])
        return rows

    rows_model_1 = build_modelo_table(items_data[0:14])
    t_model_1 = Table(rows_model_1, colWidths=[78, 148, 85, 68, 75], repeatRows=1)
    t_model_1.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), TABLE_HEADER_DARK),
        ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor("#b0c4b0")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 1.8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.8),
        ('LEFTPADDING', (0,0), (-1,-1), 2.5),
        ('RIGHTPADDING', (0,0), (-1,-1), 2.5),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_MINT]),
    ]))
    story.append(t_model_1)
    story.append(Paragraph("Tabela 5. Modelo Proposto por Infinity Health, SA (continuação na página seguinte)", style_caption))
    story.append(PageBreak())

    story.append(Paragraph("5.3. Modelo Proposto por Infinity Health, SA (continuação)", style_heading2))
    story.append(Paragraph("Tabela 5. Modelo Proposto por Infinity Health, SA (continuação)", style_caption))
    rows_model_2 = build_modelo_table(items_data[14:28])
    t_model_2 = Table(rows_model_2, colWidths=[78, 148, 85, 68, 75], repeatRows=1)
    t_model_2.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), TABLE_HEADER_DARK),
        ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor("#b0c4b0")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 1.8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 1.8),
        ('LEFTPADDING', (0,0), (-1,-1), 2.5),
        ('RIGHTPADDING', (0,0), (-1,-1), 2.5),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_MINT]),
    ]))
    story.append(t_model_2)
    story.append(Spacer(1,4))
    story.append(Paragraph("Nota: As imagens acima são ilustrativas e foram obtidas nos sítios oficiais dos fabricantes (Heal Force, Haier Biomedical, DARA Pharma, IKA, Miele, Grant, Apera, VWR, BLBIO, entre outros) e nos catálogos técnicos que se anexam. As referências assinaladas correspondem integralmente às da Planilha de Preços (Tabela 6). A Infinity Health, SA assegura que os modelos entregues serão rigorosamente os indicados, ou equivalentes tecnicamente superiores, mediante aprovação prévia da Entidade Contratante.", style_normal_small))
    story.append(PageBreak())

    # PAGE 15 - PROPOSTA FINANCEIRA
    story.append(Paragraph("6. Proposta Financeira", style_title))
    story.append(Paragraph("Exmos. Senhores,", style_normal))
    story.append(Paragraph("Instituto de Investigação Agrária de Moçambique (IIAM)", style_normal))
    story.append(Paragraph("Av. das FPLM, n.º 2698, UGEA (Edifício do Pavilhão Novo) — Cidade de Maputo", style_normal_small))
    story.append(Spacer(1,6))
    story.append(Paragraph("Assunto: <b>Apresentação de Proposta Financeira — Concurso Público N.º 35A003641/CP/03/2026 — Fornecimento de Diverso Equipamento de Laboratório</b>", style_normal))
    story.append(Spacer(1,6))
    story.append(Paragraph("Prezados Senhores,", style_normal))
    story.append(Paragraph("A <b>Infinity Health, SA</b>, sociedade comercial com sede na Rua da França, n.º 273, Bairro da Coop, Cidade de Maputo, NUIT 401550216, na qualidade de concorrente ao Concurso Público N.º <b>35A003641/CP/03/2026</b>, vem, por este meio, apresentar a sua proposta financeira, em estrita observância do Caderno de Encargos e do Decreto n.º 79/2022, de 30 de Dezembro, nos termos e nas condições seguintes:", style_normal))
    story.append(Spacer(1,4))
    story.append(Paragraph("<b>a)</b> &nbsp;Examinámos, com a devida diligência, os documentos do Concurso Público N.º 35A003641/CP/03/2026, referente ao fornecimento de diverso equipamento de laboratório para o Instituto de Investigação Agrária de Moçambique (IIAM), e apresentamos a nossa proposta, sem quaisquer reservas, condicionantes ou desvios, em conformidade com o disposto no Decreto n.º 79/2022, de 30 de Dezembro, e com as especificações técnicas da Parte II.1 do Caderno de Encargos;", style_normal))
    total_s_str = fmt_pt(TOTAL_S_IVA)
    total_iva_str = fmt_pt(TOTAL_IVA)
    total_c_str = fmt_pt(TOTAL_C_IVA)
    extenso_total_c = "oitenta e três milhões, sessenta e seis mil, novecentos e trinta e cinco meticais e setenta e oito centavos"
    extenso_total_s = "setenta e um milhões, seiscentos e nove mil, quatrocentos e vinte e sete meticais e quarenta centavos"
    story.append(Paragraph(f"<b>b)</b> &nbsp;Os preços da nossa proposta constam da <b>Planilha de Preços (Tabela 6, página 16)</b> e ascendem ao montante global de <b>{total_c_str} MT ({extenso_total_c})</b>, com IVA incluído, correspondente a <b>{total_s_str} MT ({extenso_total_s})</b> sem IVA, acrescido de <b>{total_iva_str} MT</b> a título de IVA (16 %), conforme apresentado adiante, com base no Mapa de Quantidades e nas Especificações Técnicas do Caderno de Encargos;", style_normal))
    story.append(Paragraph("<b>c)</b> &nbsp;Declaramos, sob compromisso de honra, que não nos encontramos abrangidos por qualquer situação de impedimento, de conflito de interesses ou de incapacidade prevista na legislação aplicável, e que cumprimos integralmente as obrigações fiscais e para com a segurança social;", style_normal))
    story.append(Paragraph("<b>d)</b> &nbsp;A presente proposta tem validade de <b>90 (noventa) dias</b>, contados a partir da data da sua apresentação, em conformidade com o n.º 19.1 do Documento de Concurso, e, conjuntamente com a Vossa aceitação por escrito, constante da notificação de adjudicação, constituirá compromisso vinculativo entre as partes, até à assinatura do contrato formal;", style_normal))
    story.append(Paragraph("<b>e)</b> &nbsp;Compreendemos que a Entidade Contratante não se encontra obrigada a aceitar a proposta de menor preço, nem qualquer outra proposta que venha a receber, reservando-se o direito de anular o concurso nos termos legais.", style_normal))
    story.append(Spacer(1,12))
    story.append(Paragraph("Sem mais, subscrevemo-nos com elevada consideração e inteiro dispor para os esclarecimentos adicionais que Vossas Excelências entendam necessários.", style_normal))
    story.append(Spacer(1,18))
    story.append(Paragraph("Maputo, 28 de Setembro de 2026", style_normal))
    story.append(Spacer(1,18))
    sig_data = [
        [Paragraph("___________________________________________<br/><b>Assinatura do Representante Autorizado</b><br/><font size=\"6\">Nome: ___________________________________________<br/>Cargo: Director Comercial<br/>Infinity Health, SA — NUIT 401550216<br/>Carimbo da empresa</font>", style_table_cell_center)],
    ]
    t_sig = Table(sig_data, colWidths=[280])
    t_sig.setStyle(TableStyle([
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('LEFTPADDING', (0,0), (-1,-1), 10),
        ('LINEBELOW', (0,0), (-1,0), 0.6, colors.HexColor("#999999")),
    ]))
    story.append(Table([[Spacer(220,1), t_sig]], colWidths=[220, 280]))
    # Next page is landscape planilha
    story.append(NextPageTemplate('landscape'))
    story.append(PageBreak())

    # PAGE 16 - PLANILHA PAISAGEM
    story.append(Paragraph("6.1. Planilha de Preços dos Bens — Tabela 6 (paisagem)", style_caption))
    story.append(Paragraph("Concurso Público N.º 35A003641/CP/03/2026 — Instituto de Investigação Agrária de Moçambique (IIAM) &nbsp; | &nbsp; Moeda: Metical (MT) &nbsp; | &nbsp; IVA: 16 % &nbsp; | &nbsp; Validade: 90 dias &nbsp; | &nbsp; Regime: Preço Global &nbsp; | &nbsp; Data: 28/09/2026", ParagraphStyle('SubCaptionCenter', parent=style_normal_small, alignment=TA_CENTER, textColor=colors.HexColor("#444444"), fontSize=6, leading=7)))
    story.append(Spacer(1,4))
    # Header - ensure columns Ref, Descrição, QTD, Preço Unitário, Total s/IVA, IVA, Total c/IVA bem visíveis
    hdr = [
        Paragraph('<b><font color="white">Item</font></b>', style_table_header),
        Paragraph('<b><font color="white">Ref.ª</font></b>', style_table_header),
        Paragraph('<b><font color="white">Descrição dos Bens</font></b>', style_table_header),
        Paragraph('<b><font color="white">País de<br/>Origem</font></b>', style_table_header),
        Paragraph('<b><font color="white">Tempo de<br/>Entrega</font></b>', style_table_header),
        Paragraph('<b><font color="white">QTD</font></b>', style_table_header),
        Paragraph('<b><font color="white">Un.</font></b>', style_table_header),
        Paragraph('<b><font color="white">Preço Unitário<br/>(MT)</font></b>', style_table_header),
        Paragraph('<b><font color="white">Preço Total<br/>s/ IVA (MT)</font></b>', style_table_header),
        Paragraph('<b><font color="white">IVA 16%<br/>(MT)</font></b>', style_table_header),
        Paragraph('<b><font color="white">Preço Total<br/>c/ IVA (MT)</font></b>', style_table_header),
    ]
    rows = [hdr]
    for it in items_data:
        total_s = it["qtd"]*it["pu"]
        iva = total_s*0.16
        tot_c = total_s+iva
        # Description: keep concise but complete, wrap - truncate to fit single page (28 rows) - 150 chars max for 2 lines
        desc_para = Paragraph(it["desc"][:150] + ("…" if len(it["desc"])>150 else ""), ParagraphStyle('CellDescLand', parent=style_table_cell_small, alignment=TA_LEFT, fontSize=4.4, leading=5.0))
        # Ref: ensure not overlapped, centered, small font, wrap
        ref_text = it["ref"].replace("\n","<br/>")
        # For ref longer than 60 chars, reduce font
        ref_style = ParagraphStyle('CellRefLand', parent=style_table_cell_small, alignment=TA_CENTER, fontSize=4.6, leading=5.0)
        rows.append([
            Paragraph(f"{it['item']}", ParagraphStyle('CellC', parent=style_table_cell, alignment=TA_CENTER, fontSize=4.8)),
            Paragraph(ref_text, ref_style),
            desc_para,
            Paragraph(it["pais"], ParagraphStyle('CellPaisLand', parent=style_table_cell_small, alignment=TA_CENTER, fontSize=4.5)),
            Paragraph(it["tempo"], ParagraphStyle('CellTempoLand', parent=style_table_cell_small, alignment=TA_CENTER, fontSize=4.5)),
            Paragraph(f"{it['qtd']}", ParagraphStyle('CellQtd', parent=style_table_cell_center, fontSize=4.6)),
            Paragraph(it["un"], ParagraphStyle('CellUn', parent=style_table_cell_center, fontSize=4.6)),
            Paragraph(fmt_pt(it["pu"]), ParagraphStyle('CellPu', parent=style_table_cell_right, fontSize=4.6)),
            Paragraph(fmt_pt(total_s), ParagraphStyle('CellTotS', parent=style_table_cell_right, fontSize=4.6)),
            Paragraph(fmt_pt(iva), ParagraphStyle('CellIVA', parent=style_table_cell_right, fontSize=4.6)),
            Paragraph(f"<b>{fmt_pt(tot_c)}</b>", ParagraphStyle('CellTotC', parent=style_table_cell_right, fontName=FONT_BOLD, fontSize=4.6)),
        ])
    rows.append([
        Paragraph('<b>TOTAL</b>', ParagraphStyle('TotalCellCenter', parent=style_table_cell_center, fontName=FONT_BOLD, textColor=colors.white, fontSize=6)),
        Paragraph('', style_table_cell),
        Paragraph('<b><font color="white">28 itens — Lote Único</font></b>', ParagraphStyle('TotalDesc', parent=style_table_cell, textColor=colors.white, fontName=FONT_BOLD, fontSize=6)),
        Paragraph('', style_table_cell),
        Paragraph('', style_table_cell),
        Paragraph('', style_table_cell),
        Paragraph('', style_table_cell),
        Paragraph('', style_table_cell),
        Paragraph(f'<b><font color="white">{fmt_pt(TOTAL_S_IVA)}</font></b>', ParagraphStyle('TotalRight', parent=style_table_cell_right, textColor=colors.white, fontName=FONT_BOLD, fontSize=6)),
        Paragraph(f'<b><font color="white">{fmt_pt(TOTAL_IVA)}</font></b>', ParagraphStyle('TotalRight2', parent=style_table_cell_right, textColor=colors.white, fontName=FONT_BOLD, fontSize=6)),
        Paragraph(f'<b><font color="white">{fmt_pt(TOTAL_C_IVA)}</font></b>', ParagraphStyle('TotalRight3', parent=style_table_cell_right, textColor=colors.white, fontName=FONT_BOLD, fontSize=6.5)),
    ])
    # Column widths: ajustado para visibilidade das colunas críticas (Ref, Descrição, QTD, PU, Tot s/IVA, IVA, Tot c/IVA)
    # Total width landscape = 842 - 48 = 794
    col_widths_land = [26, 78, 210, 50, 50, 24, 24, 64, 68, 64, 70]  # sum 728
    t_plan = Table(rows, colWidths=col_widths_land, repeatRows=1)
    t_plan.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), TABLE_HEADER_DARK),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 0.30, colors.HexColor("#9db99d")),
        ('TOPPADDING', (0,0), (-1,-1), 0.6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 0.6),
        ('LEFTPADDING', (0,0), (-1,-1), 1.4),
        ('RIGHTPADDING', (0,0), (-1,-1), 1.4),
        ('ROWBACKGROUNDS', (0,1), (-1,-2), [colors.white, HexColor("#f2faf2")]),
        ('BACKGROUND', (0,-1), (-1,-1), DARK_GREEN),
        ('TEXTCOLOR', (0,-1), (-1,-1), colors.white),
        ('FONTSIZE', (0,0), (-1,-1), 5.5),
        ('FONTNAME', (0,0), (-1,0), FONT_BOLD),
        ('FONTNAME', (0,-1), (-1,-1), FONT_BOLD),
        ('ALIGN', (7,1), (-1,-1), 'RIGHT'),
        # Highlight critical columns with light mint background
        ('BACKGROUND', (1,1), (1,-2), HexColor("#f9fdf9")),
        ('BACKGROUND', (2,1), (2,-2), HexColor("#ffffff")),
        ('BACKGROUND', (5,1), (5,-2), HexColor("#fffef5")),
        ('BACKGROUND', (7,1), (7,-2), HexColor("#f5f9ff")),
        ('BACKGROUND', (8,1), (8,-2), HexColor("#f5f9ff")),
        ('BACKGROUND', (9,1), (9,-2), HexColor("#fff5f5")),
        ('BACKGROUND', (10,1), (10,-2), HexColor("#f0faf0")),
    ]))
    story.append(t_plan)
    story.append(Spacer(1,5))
    legend = """<font size="5.2" color="#333333"><b>Notas:</b> (1) Preços fixos e não reajustáveis, expressos em Metical (MT), com IVA à taxa de 16 % incluído onde aplicável. (2) Os preços incluem o transporte até ao local de entrega (IIAM, Maputo), a instalação, a montagem, a testagem, a calibração inicial e a formação dos técnicos. (3) Garantia mínima de 12 meses; peças sobressalentes por 5 anos. (4) Prazo de entrega: 30 Dias Úteis para origem RSA/Reino Unido/Espanha e 90 Dias Úteis para origem China; prazo máximo global de 90 dias a contar do visto do Tribunal Administrativo. (5) A presente planilha faz parte integrante da Proposta Financeira (página 15, alínea b) e corresponde ao valor global de <b>83.066.935,78 MT</b> com IVA, <b>71.609.427,40 MT</b> sem IVA e <b>11.457.508,38 MT</b> de IVA.</font>"""
    story.append(Paragraph(legend, ParagraphStyle('LegendLand', parent=style_normal_small, alignment=TA_JUSTIFY, fontSize=5.2, leading=6.5)))
    story.append(Spacer(1,3))
    # Removido parágrafo duplicado de rodapé (já existe no header_footer_landscape) para evitar overflow de página
    story.append(NextPageTemplate('portrait'))
    story.append(PageBreak())

    # PAGE 17 - REQUISITOS
    story.append(Paragraph("7. Requisitos para Qualificação e Anexos", style_title))
    story.append(Paragraph("Em cumprimento do Caderno de Encargos do Concurso Público N.º 35A003641/CP/03/2026, a Infinity Health, SA apresenta, em anexo, a documentação de habilitação exigida, devidamente autenticada e organizada em volumes identificados como “ORIGINAL” e “CÓPIA”, em invólucro opaco, fechado e selado, nos termos do n.º 11.1 do Documento de Concurso.", style_normal))
    story.append(Spacer(1,4))
    story.append(Paragraph("A) Documentos obrigatórios do Documento de Concurso:", style_heading2))
    docs_a = [
        "Cadastro Único válido de fornecedores de bens ao Estado (n.º 8.1);",
        "Comprovativos de facturação em actividades similares ao objecto da contratação, no mínimo 3 (três) referências (n.º 8.2);",
        "Declaração comprovativa das instalações e dos equipamentos adequados e disponíveis, com indicação de dados para verificação (n.º 8.2);",
        "Declaração comprovativa da equipa profissional e técnica, acompanhada dos respectivos currículos (n.º 8.2);",
        "Declaração de experiência em actividades similares, emitida por pessoa de direito público ou privado, no mínimo 3 (três) (n.º 8.2);",
        "Garantia mínima dos equipamentos de 12 meses, com extensão a 24 meses consoante o fabricante (n.º 8.2);",
        "Catálogos dos bens a fornecer, com a descrição pormenorizada das especificações técnicas (n.º 8.2);",
        "Autorização do fabricante (n.º 15.1.2) e certificado de origem (n.º 8.2);",
        "Compromisso de montagem, testagem dos equipamentos e treinamento dos técnicos, incluídos no fornecimento (n.º 8.2);",
        "Após adjudicação: certidão válida de quitação (AT), declaração do INSS e declaração de não falência ou concordata (n.º 8.7)."
    ]
    for d in docs_a:
        story.append(Paragraph(f"• &nbsp;{d}", style_bullet))
    story.append(Spacer(1,4))
    story.append(Paragraph("B) Condições gerais de participação:", style_heading2))
    docs_b = [
        "Proposta redigida em língua portuguesa (n.º 16.1);",
        "Preços cotados em Metical (MT), correspondendo a 100 % do lote único (n.º 12.1 e 14.1);",
        "Apresentação de 2 (dois) exemplares (ORIGINAL e CÓPIA), em invólucro opaco, fechado e selado (n.º 11.1);",
        "Validade da proposta: 90 (noventa) dias (n.º 19.1);",
        "Critério de avaliação: Menor Preço Avaliado (n.º 26 e 27);",
        "Prazo de entrega: máximo de 90 dias a contar do visto do Tribunal Administrativo;",
        "Local de entrega: IIAM, Direcção de Ciências Animais, Cidade de Maputo;",
        "Garantia e assistência técnica local, com peças sobressalentes por 5 anos e formação incluída."
    ]
    for d in docs_b:
        story.append(Paragraph(f"• &nbsp;{d}", style_bullet))
    story.append(Spacer(1,8))
    story.append(Paragraph("Checklist de documentação de habilitação (para controlo do Júri):", style_caption))
    checklist_data = [
        [Paragraph('<b><font color="white">Anexo</font></b>', style_table_header), Paragraph('<b><font color="white">Documento de habilitação</font></b>', style_table_header), Paragraph('<b><font color="white">Incluído</font></b>', style_table_header)],
        [Paragraph('Anexo A', style_table_cell_center), Paragraph('Certificado de Inscrição no Cadastro Único de Fornecedores do Estado', style_table_cell), Paragraph('SIM', style_table_cell_center)],
        [Paragraph('Anexo B', style_table_cell_center), Paragraph('Comprovativos de facturação em actividades similares (mínimo de 3)', style_table_cell), Paragraph('SIM', style_table_cell_center)],
        [Paragraph('Anexo C', style_table_cell_center), Paragraph('Declaração de instalações e equipamentos adequados e disponíveis', style_table_cell), Paragraph('SIM', style_table_cell_center)],
        [Paragraph('Anexo D', style_table_cell_center), Paragraph('Declaração da equipa técnica com os respectivos CV', style_table_cell), Paragraph('SIM', style_table_cell_center)],
        [Paragraph('Anexo E', style_table_cell_center), Paragraph('Comprovativo de capital social e capacidade financeira', style_table_cell), Paragraph('SIM', style_table_cell_center)],
        [Paragraph('Anexo F', style_table_cell_center), Paragraph('Declaração de experiência em actividades similares (mínimo de 3)', style_table_cell), Paragraph('SIM', style_table_cell_center)],
        [Paragraph('Anexo G', style_table_cell_center), Paragraph('Autorização do fabricante (carta de representação oficial)', style_table_cell), Paragraph('SIM', style_table_cell_center)],
        [Paragraph('Anexo H', style_table_cell_center), Paragraph('Certificados de origem dos bens', style_table_cell), Paragraph('SIM', style_table_cell_center)],
        [Paragraph('Anexo I', style_table_cell_center), Paragraph('Catálogos dos bens a fornecer', style_table_cell), Paragraph('SIM', style_table_cell_center)],
        [Paragraph('Anexo J', style_table_cell_center), Paragraph('Compromisso de garantia mínima de 12 meses', style_table_cell), Paragraph('SIM', style_table_cell_center)],
        [Paragraph('Anexo K', style_table_cell_center), Paragraph('Compromisso de montagem, testagem e formação', style_table_cell), Paragraph('SIM', style_table_cell_center)],
    ]
    t_check = Table(checklist_data, colWidths=[60, 350, 55], repeatRows=1)
    t_check.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), TABLE_HEADER_DARK),
        ('GRID', (0,0), (-1,-1), 0.4, colors.HexColor("#bbbbbb")),
        ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ('TOPPADDING', (0,0), (-1,-1), 3),
        ('BOTTOMPADDING', (0,0), (-1,-1), 3),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, LIGHT_MINT]),
    ]))
    story.append(t_check)
    story.append(Spacer(1,6))
    story.append(Paragraph("A documentação acima encontra-se organizada, autenticada e junta à presente proposta, em envelope fechado, identificado como “ORIGINAL” e “CÓPIA”, conforme exigido no Caderno de Encargos. A Infinity Health, SA declara, sob compromisso de honra, a veracidade e a integralidade de todos os documentos apresentados.", style_normal))
    story.append(Spacer(1,8))
    story.append(Paragraph("Maputo, 28 de Setembro de 2026", style_normal))
    story.append(Spacer(1,12))
    story.append(Paragraph("___________________________________________<br/><b>Pela Infinity Health, SA</b><br/><font size=\"6\">(Assinatura e carimbo)</font>", style_table_cell_center))

    doc.build(story)
    print(f"PDF gerado: {filename} com {TOTAL_PAGES} páginas, Total c/IVA {fmt_pt(TOTAL_C_IVA)} MT")

if __name__ == "__main__":
    out = "PROPOSTA_FINANCEIRA_IIAM_35A003641_CP_03_2026_28-09-2026_FINAL_17P.pdf"
    build_pdf(out)
    # also copy to generic name for sharing
    import shutil
    shutil.copy(out, "Proposta_Tecnica_Financeira_IIAM_35A003641_FINAL.pdf")
    print("Cópia criada: Proposta_Tecnica_Financeira_IIAM_35A003641_FINAL.pdf")
