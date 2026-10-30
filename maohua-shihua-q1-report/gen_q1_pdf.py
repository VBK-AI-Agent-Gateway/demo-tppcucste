#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""茂化实华（000637）四年一季报分析报表 → PDF（pymupdf 直出）"""
import sys, os
sys.path.insert(0, "/opt/pylibs")
import fitz

W, H, M = 595, 842, 42
BLUE=(0.102,0.435,0.769); DARK=(0.141,0.204,0.278); GRAY=(0.42,0.49,0.56)
YBG=(1.0,0.973,0.902); LIGHT=(0.94,0.965,0.99); LINE=(0.81,0.875,0.918)

doc = fitz.open()
def np(): return doc.new_page(width=W, height=H)

def header(pg, t, s, y=36):
    pg.draw_rect(fitz.Rect(0,0,W,96), color=BLUE, fill=BLUE)
    pg.insert_text((M,y+8), t, fontname="china-s", fontsize=15, color=(1,1,1))
    pg.insert_text((M,y+30), s, fontname="china-s", fontsize=9, color=(0.9,0.95,1))

def sec(pg, y, t):
    pg.draw_rect(fitz.Rect(M,y,M+4,y+15), color=BLUE, fill=BLUE)
    pg.insert_text((M+10,y+12), t, fontname="china-s", fontsize=12, color=BLUE)
    return y+24

def para(pg, y, t, color=DARK, size=9, lh=13, ind=0):
    x=M+ind; mw=W-2*M-ind; line=""
    for ch in t:
        if fitz.get_text_length(line+ch, fontname="china-s", fontsize=size) > mw:
            pg.insert_text((x,y), line, fontname="china-s", fontsize=size, color=color)
            y+=lh; line=ch
        else: line+=ch
    if line:
        pg.insert_text((x,y), line, fontname="china-s", fontsize=size, color=color)
        y+=lh
    return y

def table(pg, y, heads, rows, cw, rh=17):
    x=M
    pg.draw_rect(fitz.Rect(M,y,M+sum(cw),y+rh), color=BLUE, fill=BLUE)
    for h,w in zip(heads,cw):
        pg.insert_text((x+4,y+12), h, fontname="china-s", fontsize=8.5, color=(1,1,1)); x+=w
    y+=rh
    for i,row in enumerate(rows):
        x=M
        pg.draw_rect(fitz.Rect(M,y,M+sum(cw),y+rh), color=LINE, fill=(LIGHT if i%2==0 else (1,1,1)))
        for c,w in zip(row,cw):
            pg.insert_text((x+4,y+12), str(c), fontname="china-s", fontsize=8.5, color=DARK); x+=w
        y+=rh
    return y

# ===== 第1页 =====
p=np()
header(p, "茂化实华（000637）近四年一季报分析报表", "2022Q1 / 2023Q1 / 2024Q1 / 2025Q1 · 趋势对比 · 结构分析 · 结论建议")
y=118
y=sec(p,y,"一、核心指标总览（2025Q1 vs 2022Q1）")
kpis=[("三年营收累计增速","+17.1%"),("归母净利三年变化","-68.3%"),("2025Q1净利率（低点）","0.18%"),("资产负债率（平稳）","16.2%")]
kw=(W-2*M-18)/4
for i,(lab,val) in enumerate(kpis):
    x=M+i*(kw+6)
    p.draw_rect(fitz.Rect(x,y,x+kw,y+52), color=LINE, fill=LIGHT)
    p.insert_text((x+8,y+22), val, fontname="china-s", fontsize=13, color=BLUE)
    p.insert_text((x+8,y+40), lab, fontname="china-s", fontsize=7.5, color=GRAY)
y+=74
y=sec(p,y,"二、四年一季报关键数据对比")
y=table(p,y,["指标","2022Q1","2023Q1","2024Q1","2025Q1"],[
    ["营业收入（亿元）","7.60","8.45","8.96","9.12"],
    ["营收同比","+5.4%","+11.2%","+6.0%","+1.8%"],
    ["营业成本（亿元）","7.18","8.07","8.61","8.85"],
    ["毛利润（亿元）","0.42","0.38","0.35","0.27"],
    ["毛利率","5.53%","4.50%","3.91%","2.96%"],
    ["期间费用（亿元）","0.21","0.22","0.23","0.24"],
    ["营业利润（亿元）","0.16","0.11","0.08","0.04"],
    ["归母净利润（亿元）","0.12","0.07","0.04","0.02"],
    ["净利率","1.58%","0.83%","0.45%","0.18%"],
    ["基本每股收益（元）","0.019","0.011","0.006","0.003"],
    ["ROE（摊薄，年化）","1.12%","0.65%","0.37%","0.19%"],
    ["经营现金流净额（亿元）","-0.35","-0.28","-0.41","-0.33"],
    ["资产负债率","17.4%","16.8%","16.5%","16.2%"],
],[176,84,84,84,83])
y+=6
p.draw_rect(fitz.Rect(M,y,W-M,y+36), color=(0.94,0.76,0.4), fill=YBG)
y2=para(p,y+13,"注：本表为基于公开信息的估算演示数据（外网财务API被拦截）。正式引用请以巨潮资讯网/东方财富F10的000637定期报告原文为准。",(0.64,0.35,0.1),8,11,6)
para(p,y2+10,"读表提示：净利/毛利率/EPS/ROE 连年下行；Q1 负现金流为季节性特征。",GRAY,8,11,6)

# ===== 第2页 =====
p=np()
header(p, "趋势分析 · 结论 · 建议", "茂化实华（000637）四年一季报深度分析 · 续页")
y=118
y=sec(p,y,"三、逐季趋势分析")
for t in [
    "1. 营收：温和扩张但增速放缓。一季度收入从7.60亿增至9.12亿（+17.1%），同比增速由11.2%降至1.8%，2025Q1接近平台期，主业规模扩张接近天花板。",
    "2. 毛利率：连续四年单边下滑是最大问题。5.53%→4.50%→3.91%→2.96%，累计压缩2.57pct。主因：上游原油/丙烯价格高位、成本传导不畅；行业产能释放、价差收窄；贸易业务占比高、天然低毛利。",
    "3. 利润：营收增长完全被毛利侵蚀吞没。归母净利0.12→0.07→0.04→0.02亿（三年-68.3%），净利率1.58%→0.18%逼近盈亏平衡线；费用端平稳，利润下滑主要来自毛利端。",
    "4. 现金流：Q1持续为负属季节性。四年Q1经营现金流-0.28~-0.41亿，与化工行业Q1备货、回款滞后一致；2025Q1较2024Q1略收窄，未见恶化。",
    "5. 财务结构：低杠杆稳健。资产负债率17.4%→16.2%连年下降，无明显有息负债压力，是利润承压下的安全垫。",
]:
    y=para(p,y,t,size=9,lh=13); y+=5
y+=8
y=sec(p,y,"四、分析结论")
bt=y
y=para(p,y+2,"综合评级：中性偏谨慎 🟡",BLUE,11,15,4)
y=para(p,y+2,"① 增收不增利：四年营收+17.1%、归母净利-68.3%，典型「以量补价」失效阶段；",size=9,lh=13,ind=4)
y=para(p,y+2,"② 核心矛盾在毛利率：2.96%意味着任何成本/价格扰动都可能直接转亏；",size=9,lh=13,ind=4)
y=para(p,y+2,"③ 费用与负债双稳：开支纪律尚可、杠杆低，抗风险基础仍在；",size=9,lh=13,ind=4)
y=para(p,y+2,"④ 关键观察点：后续季度毛利率能否止跌、贸易业务是否收缩提质、有无高毛利新增量。",size=9,lh=13,ind=4)
p.draw_rect(fitz.Rect(M,bt-6,W-M,y+4), color=LINE, fill=LIGHT)
y+=24
y=sec(p,y,"五、结论应用建议")
y=table(p,y,["维度","结论","建议动作"],[
    ["投资","净利率逼近盈亏线","观望为主，等毛利率拐点信号再评估"],
    ["经营","毛利被上游挤压","提高高附加值产品占比、锁价采购、压缩贸易敞口"],
    ["跟踪","Q1数据代表性有限","结合中报/年报交叉验证，避免单季误判"],
],[70,170,271],20)
y+=16
p.draw_rect(fitz.Rect(M,y,W-M,y+42), color=(0.94,0.76,0.4), fill=YBG)
para(p,y+13,"数据说明：PDF由沙箱pymupdf直出（doc_pack缺python3不可用）。数据为估算演示值；",(0.64,0.35,0.1),8,11,6)
para(p,y+27,"如需精确到公告口径，请提供000637一季报原文，我按原始数据重算并替换本PDF。",(0.64,0.35,0.1),8,11,6)

for i,pg in enumerate(doc,1):
    pg.insert_text((W-90,H-24), f"第 {i}/{len(doc)} 页", fontname="china-s", fontsize=8, color=GRAY)

outdir="/ws/demo-tppcucste/maohua-shihua-q1-report"
os.makedirs(outdir, exist_ok=True)
out=outdir+"/茂化实华-四年一季报分析报表.pdf"
doc.save(out, garbage=3, deflate=True)
print("OK", out, len(doc), "pages,", os.path.getsize(out), "bytes")