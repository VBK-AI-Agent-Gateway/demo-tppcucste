#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""《生理学·细胞的电生理现象》教学PPT构建：P1-P20全部页面内容 + pptx输出"""
import sys, os, json
sys.path.insert(0, "/ws/.pylibs"); sys.path.insert(0, "/opt/pylibs")
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from pptx.enum.shapes import MSO_SHAPE

D = os.path.dirname(os.path.abspath(__file__))
BLUE = RGBColor(0x1A,0x6F,0xC4); DARK = RGBColor(0x24,0x34,0x47); GRAY = RGBColor(0x6B,0x7B,0x8D)
RED = RGBColor(0xD6,0x45,0x45); WHITE = RGBColor(0xFF,0xFF,0xFF); LIGHT = RGBColor(0xF2,0xF6,0xFA)
ORANGE_BG = RGBColor(0xFF,0xF8,0xE6); ORANGE = RGBColor(0xF0,0xB2,0x66); ORANGE_TX = RGBColor(0xA3,0x5A,0x1A)
NOTE_TX = RGBColor(0x7A,0x4A,0x12)

ALL = [
# ---------- P1 封面 ----------
{"page":1,"title":"封面页","type":"cover"},
# ---------- P2 目录 ----------
{"page":2,"title":"目录 · CONTENTS","type":"toc","items":[
 ["01","导入与基础","为什么学电生理 / 离子跨膜转运 / 钠钾泵 / Nernst方程"],
 ["02","静息电位 RP","定义数值 / 极化·超极化·去极化·复极化"],
 ["03","动作电位 AP","定义特征 / 时相拆解 / 离子机制 / 阈电位"],
 ["04","传导与局部电位","局部电流学说 / 跳跃式传导 / 局部电位 / 对比表"],
 ["05","兴奋性周期","绝对不应期·相对不应期·超常期·低常期 / 最高兴奋频率"],
 ["06","临床联系与考点","心电肌电 / 利多卡因 / 高低钾血症 / 易错辨析"],
 ["07","小结与作业","框架小结 / 课后思考 / 参考文献"]]},
# ---------- P3 导入 ----------
{"page":3,"title":"导入：为什么要学细胞电生理？",
 "sections":[["临床现象的共同底层",
  ["心搏、思维、随意运动——本质都是细胞膜电位变化",
   "心电图ECG、肌电图EMG、脑电图EEG = 大量细胞电活动的体表总和",
   "★考点：电生理是诊断学（心电/肌电/神经传导）的理论基石"]],
 ["一次课你要拿到的四把钥匙",
  ["静息电位怎么产生（RP机制）","动作电位怎么发生与传导（AP机制+传导）",
   "局部电位是什么、和AP差在哪","兴奋性如何随时间周期性变化"]],
 ["衔接临床",["麻醉、抗心律失常、电解质紊乱——全部回归到本节机制"]]],
 "note":"开场提问：「为什么心脏起搏点能自动节律性跳动？为什么局麻药能止痛？」先抛问题留悬念，本节课结束时回收答案。强调：这节课听懂，后面诊断学心电图、麻醉学局麻药就通了一半。",
 "fig":"三联卡片式示意图：左「心脏搏动」中「神经信号」右「骨骼肌收缩」，底部汇聚到蓝色图标「细胞膜电位变化」，向下发散为ECG/EMG/EEG三条波形曲线。"},
# ---------- P4 前置基础 ----------
{"page":4,"title":"前置基础：离子跨膜转运 · 钠钾泵 · 离子平衡电位",
 "sections":[["① 细胞膜离子跨膜转运（回顾）",
  ["被动转运：顺浓度差，经通道/载体，不耗能（易化扩散）",
   "主动转运：逆浓度差，需耗能——钠钾泵（Na⁺-K⁺-ATPase）"]],
 ["② 钠钾泵（★熟悉）",
  ["每分解1个ATP：泵出3个Na⁺、泵入2个K⁺ → 生电性（净外向电流）",
   "建立并维持膜内外Na⁺/K⁺浓度差（胞外高Na⁺、胞内高K⁺）",
   "★考点：浓度差是「势能库」，AP/RP的能量根源"]],
 ["③ 离子平衡电位（Nernst方程·只讲意义）",
  ["意义：某离子在浓度差（外推力）与电位差（回拉力）平衡时的跨膜电位",
   "E=RT/zF·ln(Co/Ci)——本科只记：浓度决定平衡电位，不推导",
   "例：K⁺平衡电位Eₖ≈-90mV，与RP(-70~-90mV)接近 → RP主要由K⁺决定"]]],
 "note":"钠钾泵用「电梯」比喻：3出2入像每次上下楼都净送一个人到楼下（生电）。Nernst方程板书只写一行式子+一句结论：「某种离子的平衡电位，只取决于膜两侧该离子的浓度比」，不做数学推导，防止学生畏难。",
 "fig":"左：细胞膜双层+通道蛋白示意，标Na⁺(胞外高)/K⁺(胞内高)浓度箭头；中：钠钾泵卡通，3个Na⁺箭头向外、2个K⁺向内，旁挂ATP→ADP；右：天平图——浓度差(上箭头)与电位差(下箭头)平衡，指针指Eₖ=-90mV。"},
# ---------- P5 静息电位 ----------
{"page":5,"title":"静息电位 Resting Potential：定义 · 数值 · 机制",
 "sections":[["定义与数值",
  ["定义：细胞未受刺激时，膜两侧外正内负的电位差",
   "数值：神经/骨骼肌细胞约-70～-90mV（哺乳动物骨骼肌约-90mV）",
   "★考点：「安静状态、外正内负、K⁺背景电位」三要素背诵"]],
 ["产生条件（先条件→离子流→电位）",
  ["条件①：膜两侧K⁺/Na⁺浓度差（钠钾泵维持）",
   "条件②：安静时膜对K⁺通透大（K⁺漏通道开放）",
   "离子流：K⁺沿浓度差外流 → 正电荷外积 → 膜内负电位升高",
   "电位变化：外流受电位差阻挡，达Eₖ时净电流=0 → 稳定于Eₖ附近"]],
 ["特点与略偏离Eₖ的原因",
  ["特点：稳定、不波动；是AP发生的「起跑线」",
   "RP略高于Eₖ：安静时少量Na⁺内流 + 钠钾泵生电性贡献"]]],
 "note":"严格按「条件→离子流动→电位变化」三步讲，语速放慢。板书画时间轴：K⁺外流箭头逐步变短，直到被电位差完全抵消——「拔河平衡」比喻。强调负号：膜内相对膜外为负，记的是膜内电位。",
 "fig":"中央细胞剖面图：膜外标「+」、膜内标「−」；K⁺通过漏通道向外扩散(粗箭头)，Na⁺漏入(细虚线)；右侧微电位计指针显示-70mV；底部小图：钠钾泵维持浓度梯度(Na⁺向外、K⁺向内循环箭头)。"},
# ---------- P6 四概念辨析 ----------
{"page":6,"title":"静息电位补充：极化 · 超极化 · 去极化 · 复极化",
 "sections":[["四个概念（★考点·必背方向）",
  ["极化 Polarization：静息时外正内负的稳态（基线，-70mV）",
   "去极化 Depolarization：膜内负电位减小（-70→0），向0靠近",
   "反极化/超射 Overshoot：膜内转为正电位（0→+30mV）",
   "复极化 Repolarization：从峰值恢复到外正内负（+30→-70mV）",
   "超极化 Hyperpolarization：膜内负电位比静息更负（-70→-90mV）"]],
 ["方向速记（★易错）",
  ["「去」=去掉负电（向0走）；「超」=超过静息（更负）",
   "去极化/复极化的参照点都是静息电位水平",
   "后电位阶段的去极化称「负后电位」、超极化称「正后电位」（结合P8）"]]],
 "note":"让学生用手比划：手掌平放=极化；抬起(负值变小)=去极化；翻正=反极化；放回=复极化；压低=超极化。课堂快问快答：「-90mV相比-70mV是?」(超极化)——当堂检测记忆。",
 "fig":"坐标系示意图：纵轴电位(mV)、横轴时间；静息基线-70mV(标「极化」)，上升段标「去极化」，0以上标「反极化/超射」，下降段标「复极化」，下凹到-90mV处标「超极化」，不同颜色箭头标注方向。"},
# ---------- P7 AP定义特征 ----------
{"page":7,"title":"动作电位 Action Potential：定义 · 特征",
 "sections":[["定义",
  ["细胞受阈刺激/阈上刺激后，在RP基础上产生的快速、可传播的电位波动",
   "波形：去极相→反极化→复极相（锋电位+后电位）"]],
 ["三大特征（★考点·全或无）",
  ["① 全或无：刺激达阈强度才有AP；一旦产生，幅度恒定，不随刺激增强而增大",
   "② 不衰减传导：沿膜传播时幅度、波形保持不变",
   "③ 脉冲式：锋电位之后有绝对不应期 → 两次AP互不融合，呈脉冲",
   "④ 传导靠局部电流+阈电位机制保证（P11展开）"]],
 ["与阈下刺激关系",["阈下刺激→只引起局部电位，不能产生AP（引出P12）"]]],
 "note":"「全或无」用门禁比喻：门票(刺激)够阈值门全开(AP爆发)，不够就完全不开；门一旦开了，开多大不由票多寡决定。强调临床意义：神经编码靠频率不靠幅度。",
 "fig":"经典AP双相波形图：纵轴-90~+40mV、横轴时间；标注静息(-70)、锋电位(陡段)、负后电位、正后电位；三张特征小图标(全开的门/等幅传播箭头/脉冲串)配文字标签。"},
# ---------- P8 AP时相 ----------
{"page":8,"title":"动作电位时相拆解：去极相 · 反极化 · 复极相 · 后电位",
 "sections":[["锋电位（Spike）——AP主体，可触发传导",
  ["① 去极相(上升支)：-70→0mV，Na⁺快速内流",
   "② 反极化/超射：0→+30mV，膜内变正（峰）",
   "③ 复极相(下降支)：+30→-70mV，K⁺快速外流",
   "★考点：锋电位=AP中「全或无、传导」的有效成分"]],
 ["后电位（After-potential）",
  ["负后电位：复极末期膜电位低于静息前的轻度去极化残留（≈轻度去极化状态）",
   "正后电位：短暂超极化，膜电位比静息更负，然后回基线",
   "后电位与兴奋性周期密切相关（衔接P14）"]],
 ["时相—离子对应",
  ["去极相=Na⁺内流；复极相=K⁺外流；负/正后电位=泵恢复+通道门控动力学"]]],
 "note":"板书画完整AP波形，逐段上色：红=去极相、绿=复极相、黄=后电位。提示考点：负后电位其实处于「轻度去极化」方向(所以兴奋性高)，正后电位是超极化方向(兴奋性低)——为P14埋伏笔。",
 "fig":"大幅AP波形分段上色图：①②③标号在上升、峰、下降段；④负后电位(浅红)、⑤正后电位(浅蓝)圈出；图下两列小字：左「离子：Na⁺内流→K⁺外流」右「电位方向：去极→反极→复极→后电位」。"},
# ---------- P9 离子机制 ----------
{"page":9,"title":"动作电位离子机制：电压门控Na⁺/K⁺通道开闭时序",
 "sections":[["电压门控Na⁺通道——三态（★考点）",
  ["静息态(关)→激活态(开，快速内流)→失活态(关，不可即开)",
   "去极化到阈电位：激活门(m)开放→Na⁺爆发性内流→再生性正反馈",
   "达峰后：失活门(h)关闭→Na⁺通道失活→内流停止(必然复极)"]],
 ["电压门控K⁺通道——慢开慢关",
  ["去极化开始后延迟开放(慢)，复极期才充分开放",
   "K⁺外流使膜电位回落；复极完成后通道缓慢关闭",
   "★考点：Na⁺通道「快开快关带失活」、K⁺通道「慢开慢关不失活」"]],
 ["时序总结",
  ["阈电位→Na⁺通道开→Na⁺内流(陡升)→Na⁺失活+K⁺开→K⁺外流(下降)→K⁺通道缓关→回RP"]]],
 "note":"用两把「门」讲：Na⁺门是自动门禁——开了之后1ms就锁死(失活)，锁死后必须等复极才解锁；K⁺门是弹簧门——推一下慢慢开、慢慢关。再生性正反馈强调「正反馈而非负反馈」是AP爆发的原因。",
 "fig":"上下对照时序图：上=AP波形；中=Na⁺通道状态条(静息→激活→失活三色带)；下=K⁺通道状态条(关→延迟开→缓关)；竖虚线对齐「阈电位点」「峰电位点」「复极完成点」，箭头标Na⁺内流(红)/K⁺外流(绿)。"},
# ---------- P10 阈电位 ----------
{"page":10,"title":"阈电位 · 阈刺激 · 阈强度：三者逻辑关系",
 "sections":[["三概念定义（★考点·易混）",
  ["阈刺激/阈强度 Threshold stimulus：能引起AP的最小刺激强度(针对「刺激」)",
   "阈下刺激：低于阈强度→仅局部电位，不能产生AP",
   "阈电位 Threshold potential：能触发AP的膜电位临界值(针对「膜电位」，约-55mV)"]],
 ["逻辑链（背诵顺序）",
  ["阈刺激/阈上刺激 → 膜去极化达到阈电位 → Na⁺通道大量开放 → 再生性正反馈 → AP",
   "★考点：刺激作用于「膜外」，阈电位是「膜内」临界值；刺激是条件，阈电位是机制节点",
   "刺激强度↑→RP到阈电位距离缩短→更易达到→兴奋性↑(与P14衔接)"]],
 ["常考判断",
  ["「只要刺激足够强，AP幅度就更大」→错(全或无)",
   "「阈下刺激不引起任何电位变化」→错(有局部电位)"]]],
 "note":"画「起跑线-终点线」：静息是起跑线，阈电位是终点线，刺激强度决定能不能跑到终点；跑没跑过线决定AP产生与否，与跑多快(刺激再大)无关。当堂出2道判断题现场举手作答。",
 "fig":"三栏关系图：左「刺激轴」(阈下/阈/阈上三档箭头)、中「膜电位轴」(-70→-55阈电位虚线→0→+30)、右「结果」(局部电位/AP)；红色虚线横贯标注「阈电位=-55mV」；底部流程条：阈刺激→达阈电位→Na⁺通道爆发→AP。"},
# ---------- P11 传导 ----------
{"page":11,"title":"动作电位的传导：局部电流学说 · 跳跃式传导",
 "sections":[["① 局部电流学说（无髓纤维）",
  ["已兴奋段(内正外负)与邻近静息段(外正外负)间存在电位差",
   "形成局部电流→邻近膜去极化达阈电位→产生新AP",
   "★考点：AP「不衰减传导」的原因——每次都有RP能量补充，非消耗传递"]],
 ["② 有髓神经纤维·跳跃式传导(Saltatory)",
  ["髓鞘处绝缘、电容小→电流只能在郎飞结(Node of Ranvier)间跳跃",
   "AP只在郎飞结处产生→跳跃传导",
   "★考点：意义=传导速度↑(快100倍级)、能量消耗↓(少泵Na⁺)"]],
 ["对比小结",["无髓：连续性传导(一格一格)；有髓：跳跃式传导(跨格跳)"]]],
 "note":"走廊学生手拉手传「去极化」的模拟：相邻同学手递手=局部电流；跳格子=有髓跳跃。追问：「为什么局部电流能让邻近膜达到阈电位？」——因为局部电流本质是「电刺激」，强度足以触发AP。",
 "fig":"左图：一段神经纤维，兴奋段(红，膜内+)与静息段(蓝，膜外+)间画弧形局部电流箭头，右侧显示下一个AP即将产生；右图：有髓纤维纵剖，蓝色髓鞘包裹、郎飞结裸露，红色AP脉冲只在结处出现，虚线弧形箭头标「跳跃」，旁注「速度快、耗能少」。"},
# ---------- P12 局部电位 ----------
{"page":12,"title":"局部电位 Local Potential（电紧张电位）",
 "sections":[["产生条件",
  ["阈下刺激直接作用于膜→部分电压门控Na⁺通道开放→少量Na⁺内流",
   "不足以达到阈电位→不能产生AP"]],
 ["四大特点（★考点·与AP对比的基础）",
  ["① 不具有「全或无」：随刺激强度增大而增大（等级性）",
   "② 可衰减传导：以电紧张形式向四周扩布，随距离衰减",
   "③ 可叠加/总和：时间总和(连续刺激)、空间总和(多点同时)→可达阈电位",
   "④ 没有不应期：可产生多个幅度叠加"]],
 ["引申",["总和达到阈电位→即可触发AP（局部电位是「AP的引信」）"]]],
 "note":"用「涟漪」比喻局部电位：石子(阈下刺激)扔进水里起小波纹，越近越大、越远越小(衰减)；多颗石子同时扔，波纹叠加起来就能掀大浪(总和→AP)。强调局部电位是分级的、AP是全或无的。",
 "fig":"水波纹示意图：中心小石子(阈下刺激)入水，同心圆波纹向外扩散并逐渐变淡(标「衰减」)；右侧两小图：上方「时间总和」两波纹先后叠加、下方「空间总和」两处波纹相向叠加，叠加后越过红色虚线「阈电位」后出现AP脉冲。"},
# ---------- P13 对比表 ----------
{"page":13,"title":"重点对比表格：局部电位 VS 动作电位（核心考点）","type":"table",
 "headers":["比较项目","局部电位","动作电位"],
 "rows":[
  ["刺激强度","阈下刺激即可，随强度分级变化","需阈刺激/阈上刺激，全或无"],
  ["幅度","小，随刺激增大而增大","大，恒定不随刺激增大"],
  ["传导","电紧张扩布，随距离衰减","不衰减传导，幅度波形不变"],
  ["叠加","可时间/空间总和叠加","有绝对不应期，不可叠加"],
  ["不应期","无","有（绝对不应期）"],
  ["产生机制","部分Na⁺通道开放，少量Na⁺内流","Na⁺通道大量开放，再生性正反馈"],
  ["本质","分级电位，不是「兴奋」的标志","兴奋的标志，可触发传导"]],
 "note":"本页是核心考点页，逐行讲+让学生在笔记本上抄这张表。重点强调分水岭：局部电位「分级、可衰减、可叠加」vs AP「全或无、不衰减、有不应期」。期末考简答题「试比较局部电位与动作电位」按此表给分。",
 "fig":"两栏对比表格图：左栏浅蓝「局部电位」列、右栏浅橙「动作电位」列，中间深蓝分隔线；每行左侧项目名加粗；底部大号VS图标；关键差异词(全或无/不衰减/有不应期)红色加粗。"},
# ---------- P14 兴奋性四时期 ----------
{"page":14,"title":"细胞兴奋性的周期性变化：四个时期","type":"excitability",
 "periods":[
  ["绝对不应期 ARP","AP锋电位期间","Na⁺通道全部失活，无论多强刺激均不能产生新AP","0（最低）"],
  ["相对不应期 RRP","负后电位期","部分Na⁺通道复活，阈电位上移，阈上刺激才可能引起AP","逐渐恢复(<100%)"],
  ["超常期 SNP","正后电位早期","膜电位接近阈电位，阈下刺激也可能引起AP","稍高于正常(>100%)"],
  ["低常期 LNP","正后电位后期","膜超极化，距阈电位更远，需阈上刺激","略低于正常(<100%)"]],
 "note":"画在AP波形上逐一标注四段。绝对不应期对应锋电位——Na⁺通道全失活，这是全或无+不融合的结构基础。超常期对应正后电位(离阈电位近)，低常期对应正后电位后期(超极化，更远)——方向别讲反。",
 "fig":"AP波形+兴奋性双轨图：上方AP波形下挂四条色带(红=绝对不应期、橙=相对不应期、绿=超常期、蓝=低常期)，下方同步画「兴奋性水平曲线」，绝对不应期段为0贴底线、相对不应期回升、超常期略高于虚线正常值、低常期略低于虚线；两图时间轴严格对齐。"},
# ---------- P15 最高兴奋频率 ----------
{"page":15,"title":"兴奋性周期的考点延伸：为什么绝对不应期决定最高兴奋频率？",
 "sections":[["核心逻辑（★考点）",
  ["绝对不应期内无论多强刺激都不能产生新AP → 同一细胞两次AP必然被隔开",
   "最高兴奋频率 = 1 / (绝对不应期时长)",
   "例：ARP=3ms → 最高频率≈333次/秒；ARP=5ms → 200次/秒"]],
 ["应用解释",
  ["神经纤维ARP短→频率高(可达1000Hz)；心肌ERP长→频率低(数百次/分)",
   "★考点：AP脉冲式、不融合的根本原因=绝对不应期存在"]],
 ["易错点",["「最高兴奋频率」由绝对不应期决定，不是相对不应期"]]],
 "note":"用「相机快门」比喻：快门(绝对不应期)没打开前按多少次按钮都拍不了，两次拍照最小间隔=快门复位时间，最高连拍频率=1/快门时间。顺带点心肌ARP长→心率有上限→不易强直收缩(临床安全意义)。",
 "fig":"公式框图：大号「f_max = 1 / T_绝对不应期」居中；下方两根横向时间条对比：神经纤维(红段短，标3ms→333Hz)、心肌(红段长，标~300ms→低频)；右侧一串AP脉冲串，脉冲间隔不小于红段长度，双向箭头标「最小间隔=ARP」。"},
# ---------- P16 临床 ----------
{"page":16,"title":"临床联系：心电图/肌电图 · 利多卡因 · 高钾低钾血症","type":"clinical",
 "blocks":[
  ["① 心电图/肌电图的电生理基础",
   ["ECG：心室肌细胞AP同步总和的体表记录；QRS波≈心室去极化、T波≈复极化",
    "EMG：运动单位AP的时空总和；神经传导速度检测=测有髓纤维跳跃传导"]],
  ["② 钠通道阻滞剂·利多卡因（★考点）",
   ["结合电压门控Na⁺通道失活态——优先与失活态结合(状态依赖性)",
    "抑制Na⁺内流→降低0期去极化速度→传导减慢/阻止异位快速心律失常",
    "局麻原理：阻断神经Na⁺通道→阻断AP产生与传导→痛觉信号不能上传"]],
  ["③ 高钾/低钾血症（★高频考点）",
   ["高钾：膜外K⁺↑→浓度差↓→K⁺外流↓→RP绝对值↓(去极化)→Na⁺通道部分失活→兴奋性先高后低(严重时AP幅度↓、传导↓→心脏停搏风险)",
    "低钾：膜外K⁺↓→浓度差↑→K⁺外流↑→RP绝对值↑(超极化)→距阈电位远→兴奋性↓(但临床上因复极改变也可致心律失常)"]]],
 "note":"高钾是急症：讲「轻度高钾兴奋性一过性升高(心跳快)、重度则AP低幅度传导慢(心跳慢→停)」这个双相过程是考点。低钾对心肌影响提示学生注意与骨骼肌不同，考试常考「高钾→RP绝对值降低」判断题。",
 "fig":"三栏卡片图：左=ECG波形(P/QRS/T标注去极复极)；中=Na⁺通道三态图+利多卡因「锁住失活态门」锁头图标；右=天平图对比：高钾(膜外K⁺多、RP上移-70→-50)vs低钾(膜外K⁺少、RP下移-70→-90)，分别配兴奋性↑/↓箭头。"},
# ---------- P17 易错辨析 ----------
{"page":17,"title":"本章易错题、易混点汇总辨析","type":"mistakes",
 "items":[
  ["局部电位 vs 动作电位","局部电位=分级、可衰减、可叠加、无不应期；AP=全或无、不衰减、不可叠加、有绝对不应期。答「阈下刺激能否引起电位变化」→能(局部电位)，但不能引起兴奋"],
  ["阈刺激 vs 阈电位","阈刺激=最小刺激强度(膜外条件)；阈电位=膜内电位临界值(机制节点)。刺激是因，阈电位是中间环节，AP是果"],
  ["绝对不应期 vs 相对不应期","ARP=Na⁺通道失活、兴奋性0；RRP=部分复活、需阈上刺激。ARP对应锋电位、RRP对应负后电位"],
  ["「全或无」误解","不是「刺激弱就不产生」，而是「达不到阈就不产生，一旦产生幅度固定」——幅度与刺激强度无关"],
  ["超极化方向","超极化=比静息更负(-90)，去极化=负值变小(-50)。负后电位其实是去极化方向、正后电位是超极化方向(命名易错)"],
  ["RP与Eₖ","RP≈Eₖ但略低(绝对值)，因少量Na⁺内流+泵的生电性，不是完全等于Eₖ"]],
 "note":"这页留5分钟做「快问快答」，把六个辨析做成6张卡片翻转，学生答对得一分。历年期末考简答/选择题几乎都出自这六组，务必让学生用红笔在笔记上标星。",
 "fig":"六宫格辨析卡：每格一个「VS」标题+两行对比要点，左蓝右橙双色块；角标①-⑥编号；关键差异词(全或无/失活/超极化)红色加粗；整体风格类似「知识卡片墙」。"},
# ---------- P18 小结 ----------
{"page":18,"title":"课堂小结（极简框架版，方便记笔记）","type":"summary",
 "tree":[
  ["电生理框架","① RP(静息电位)：K⁺外流+泵维持→外正内负、-70~-90mV"],
  ["","② AP(动作电位)：阈电位→Na⁺内流(升)→Na⁺失活+K⁺外流(降)"],
  ["","③ 传导：局部电流 / 有髓跳跃式"],
  ["","④ 局部电位：阈下→分级、可衰减、可总和"],
  ["","⑤ 兴奋性周期：ARP→RRP→SNP→LNP"],
  ["临床锚点","利多卡因(阻Na⁺失活通道) / 高钾RP↓ / 低钾RP↑"]],
 "note":"让学生合上书，跟着这6行框架口头复述一遍——能复述=听懂了。强调每行都能展开成一道简答题，这页就是「背诵索引」。",
 "fig":"思维导图/树状图：中心蓝色胶囊「细胞电生理」，向右分出5个主枝(RP/AP/传导/局部电位/兴奋性周期)，每枝末端挂2-3个关键词小胶囊；底部单独一枝橙色「临床」挂利多卡因/高钾/低钾三个节点。"},
# ---------- P19 作业 ----------
{"page":19,"title":"课后思考题 + 拓展思考","type":"homework",
 "hw":[["★课后必做",
   "1. 简述静息电位的产生机制(条件→离子流→电位，8分)",
   "2. 试比较局部电位与动作电位(表格，10分)",
   "3. 为什么绝对不应期决定最高兴奋频率？(6分)",
   "4. 高钾血症患者为何易发生心律失常？(6分)"],
  ["拓展思考（选做）",
   "5. 利多卡因为什么对「正在发放冲动」的异位节律点作用更强？(状态依赖性)",
   "6. 心电图QRS波宽提示什么电生理改变？",
   "7. Hodgkin-Huxley模型如何用数学描述Na⁺/K⁺电导变化？(了解层面)"]],
 "note":"必做4题对应4个教学目标，提示按「机制三步法」作答。拓展第5题引导学有余力学生查「状态依赖性阻断」，与麻醉学衔接；第7题仅提示有兴趣同学看拓展材料，不作要求。",
 "fig":"作业清单卡片图：左侧红色五角星「★必做」区域列4题，右侧蓝色灯泡「拓展」区域列3题；每题前置序号圆点；底部画铅笔图标+「周五前提交至学习通」便签。"},
# ---------- P20 参考文献 ----------
{"page":20,"title":"参考文献页","type":"refs",
 "refs":[
  "1. 王庭槐. 生理学（第9版）. 人民卫生出版社, 2018.（第2章 细胞的基本功能）",
  "2. 朱大年, 王庭槐. 生理学学习指导与习题集（第3版）. 人民卫生出版社",
  "3. Boron WF, Boulpaep EL. Medical Physiology (3rd ed.). Elsevier, 2017. Ch.5",
  "4. Guyton AC, Hall JE. Textbook of Medical Physiology (14th ed.). Elsevier, 2020",
  "5. 中华医学会心电生理和起搏分会. 心电图标准化与解析建议",
  "6. 图片来源：自制示意图 / PhysioEx 虚拟实验 / OpenStax Anatomy & Physiology (CC-BY)"],
 "note":"提示学生教材页码直接翻到第2章第2节，课后配套习题册做「细胞基本功能」章节；英文文献供双语班/考研同学延伸。",
 "fig":"简洁文献列表页：顶部蓝色标题条，6条文献用悬挂缩进排列，左侧序号蓝色圆点；右下角小图标：书本+DOI二维码占位框；底部校徽位置留白。"},
]

# ===== 构建 PPTX =====
prs = Presentation()
prs.slide_width = Inches(13.333); prs.slide_height = Inches(7.5)
BLANK = prs.slide_layouts[6]; W = prs.slide_width

def add_rect(slide, x, y, w, h, fill, line=None, rounded=False):
    shape = slide.shapes.add_shape(MSO_SHAPE.ROUNDED_RECTANGLE if rounded else MSO_SHAPE.RECTANGLE, x, y, w, h)
    if rounded: shape.adjustments[0] = 0.05
    if fill: shape.fill.solid(); shape.fill.fore_color.rgb = fill
    else: shape.fill.background()
    if line: shape.line.color.rgb = line; shape.line.width = Pt(1)
    else: shape.line.fill.background()
    shape.shadow.inherit = False
    return shape

def add_text(slide, x, y, w, h, runs, align=PP_ALIGN.LEFT, anchor=MSO_ANCHOR.TOP, spacing=1.12):
    tb = slide.shapes.add_textbox(x, y, w, h); tf = tb.text_frame
    tf.word_wrap = True; tf.vertical_anchor = anchor
    first = True
    for item in runs:
        if isinstance(item, tuple): t, c, b, s = item
        else: t, c, b, s = item, DARK, False, 16
        p = tf.paragraphs[0] if first else tf.add_paragraph(); first = False
        p.alignment = align; p.line_spacing = spacing
        r = p.add_run(); r.text = t
        r.font.size = Pt(s); r.font.bold = b; r.font.color.rgb = c; r.font.name = "微软雅黑"
    return tb

def header(slide, no, title):
    add_rect(slide, 0, 0, W, Inches(0.8), BLUE)
    add_text(slide, Inches(0.4), Inches(0.08), Inches(12.4), Inches(0.62),
             [(title, WHITE, True, 25)], anchor=MSO_ANCHOR.MIDDLE)
    add_text(slide, Inches(12.35), Inches(6.98), Inches(0.85), Inches(0.35),
             [(f"{no}/20", GRAY, False, 11)], align=PP_ALIGN.RIGHT)

def body_with_side(slide, sections, note, figdesc):
    y = Inches(1.0)
    for head, lines in sections:
        add_text(slide, Inches(0.4), y, Inches(8.7), Inches(0.34), [(head, BLUE, True, 17)])
        y += Inches(0.375)
        for ln in lines:
            if ln.startswith("★"):
                add_text(slide, Inches(0.55), y, Inches(8.55), Inches(0.30), [(ln, RED, True, 14)])
            else:
                add_text(slide, Inches(0.55), y, Inches(8.55), Inches(0.30), [("• " + ln, DARK, False, 14)])
            y += Inches(0.325)
        y += Inches(0.11)
    rx, rw = Inches(9.2), Inches(3.75)
    fig_y, fig_h = Inches(1.0), Inches(2.3)
    add_rect(slide, rx, fig_y, rw, fig_h, LIGHT, BLUE, rounded=True)
    add_text(slide, rx+Inches(0.12), fig_y+Inches(0.06), rw-Inches(0.24), Inches(0.3),
             [("配图建议", BLUE, True, 14)])
    add_text(slide, rx+Inches(0.12), fig_y+Inches(0.40), rw-Inches(0.24), fig_h-Inches(0.48),
             [(figdesc, DARK, False, 11)], spacing=1.10)
    n_y = fig_y + fig_h + Inches(0.14); n_h = Inches(6.36) - n_y
    add_rect(slide, rx, n_y, rw, n_h, ORANGE_BG, ORANGE, rounded=True)
    add_text(slide, rx+Inches(0.12), n_y+Inches(0.06), rw-Inches(0.24), Inches(0.3),
             [("教师备注（口头讲解话术）", ORANGE_TX, True, 12)])
    add_text(slide, rx+Inches(0.12), n_y+Inches(0.38), rw-Inches(0.24), n_h-Inches(0.46),
             [(note, NOTE_TX, False, 10.5)], spacing=1.10)

def note_bar(slide, note):
    add_rect(slide, Inches(0.4), Inches(6.42), Inches(12.5), Inches(0.85), ORANGE_BG, ORANGE, rounded=True)
    add_text(slide, Inches(0.55), Inches(6.48), Inches(12.2), Inches(0.73),
             [("教师备注：" + note, NOTE_TX, False, 11)], spacing=1.08)

# P1 封面
s = prs.slides.add_slide(BLANK)
add_rect(s, 0, 0, W, prs.slide_height, BLUE)
add_text(s, Inches(1), Inches(1.7), Inches(11.3), Inches(1.6),
         [("《生理学》", RGBColor(0xCF,0xE4,0xF8), False, 30), ("细胞的电生理现象", WHITE, True, 54)], align=PP_ALIGN.CENTER)
add_text(s, Inches(1), Inches(3.5), Inches(11.3), Inches(0.5),
         [("第三章 · 第二节 · 本科临床医学专业教学课件", RGBColor(0xCF,0xE4,0xF8), False, 20)], align=PP_ALIGN.CENTER)
add_text(s, Inches(1), Inches(4.4), Inches(11.3), Inches(1.1),
         [("授课对象：临床医学大二本科生    ｜    课时：1次课 / 45分钟", WHITE, False, 18),
          ("授课人：×××    ｜    医学院生理学教研室", RGBColor(0xCF,0xE4,0xF8), False, 18)], align=PP_ALIGN.CENTER)
add_text(s, Inches(1), Inches(6.6), Inches(11.3), Inches(0.4),
         [("【第1页｜封面页】", RGBColor(0x9E,0xC8,0xEC), False, 12)], align=PP_ALIGN.CENTER)

# P2 目录
p = ALL[1]; s = prs.slides.add_slide(BLANK); header(s, 2, p["title"])
y = Inches(1.1)
for no, t, d in p["items"]:
    add_rect(s, Inches(0.5), y, Inches(12.35), Inches(0.74), LIGHT, RGBColor(0xE2,0xEC,0xF6), rounded=True)
    add_text(s, Inches(0.7), y+Inches(0.08), Inches(0.95), Inches(0.55), [(no, BLUE, True, 25)])
    add_text(s, Inches(1.75), y+Inches(0.05), Inches(3.3), Inches(0.6), [(t, DARK, True, 17)], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, Inches(5.1), y+Inches(0.08), Inches(7.6), Inches(0.55), [(d, GRAY, False, 13)], anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.83)

# 通用页 P3-P12、P15
for p in ALL:
    if p.get("type"): continue
    s = prs.slides.add_slide(BLANK); header(s, p["page"], p["title"])
    body_with_side(s, p["sections"], p["note"], p["fig"])

# P13 表格
p = next(x for x in ALL if x["page"]==13); s = prs.slides.add_slide(BLANK); header(s, 13, p["title"])
cols = [Inches(2.4), Inches(5.1), Inches(5.1)]; x0 = Inches(0.37); y = Inches(1.05); rh = Inches(0.62)
x = x0
for i, htxt in enumerate(p["headers"]):
    add_rect(s, x, y, cols[i], rh, BLUE)
    add_text(s, x, y+Inches(0.12), cols[i], Inches(0.4), [(htxt, WHITE, True, 16)], align=PP_ALIGN.CENTER)
    x += cols[i]
y += rh
for r_i, row in enumerate(p["rows"]):
    x = x0; fill = LIGHT if r_i % 2 == 0 else WHITE
    for c_i, cell in enumerate(row):
        add_rect(s, x, y, cols[c_i], Inches(0.60), fill, RGBColor(0xD8,0xE4,0xF0))
        add_text(s, x+Inches(0.1), y+Inches(0.06), cols[c_i]-Inches(0.2), Inches(0.5),
                 [(cell, DARK, c_i==0, 13)], align=PP_ALIGN.CENTER if c_i==0 else PP_ALIGN.LEFT, anchor=MSO_ANCHOR.MIDDLE, spacing=1.0)
        x += cols[c_i]
    y += Inches(0.60)
add_rect(s, Inches(0.4), Inches(5.75), Inches(12.5), Inches(0.6), LIGHT, BLUE, rounded=True)
add_text(s, Inches(0.55), Inches(5.81), Inches(12.2), Inches(0.5), [("配图建议：" + p["fig"], DARK, False, 11)], spacing=1.05)
note_bar(s, p["note"])

# P14 四时期
p = next(x for x in ALL if x["page"]==14); s = prs.slides.add_slide(BLANK); header(s, 14, p["title"])
card_w, card_h = Inches(6.15), Inches(2.1)
positions = [(Inches(0.4), Inches(1.0)), (Inches(6.75), Inches(1.0)), (Inches(0.4), Inches(3.3)), (Inches(6.75), Inches(3.3))]
colors = [RED, ORANGE, RGBColor(0x2E,0x9E,0x6B), BLUE]
for (px, py), col, (name, when, mech, exc) in zip(positions, colors, p["periods"]):
    add_rect(s, px, py, card_w, card_h, WHITE, col, rounded=True)
    add_rect(s, px, py, card_w, Inches(0.5), col, rounded=True)
    add_text(s, px+Inches(0.15), py+Inches(0.06), card_w-Inches(0.3), Inches(0.4), [(name, WHITE, True, 16)], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, px+Inches(0.15), py+Inches(0.58), card_w-Inches(0.3), card_h-Inches(0.66),
             [("对应时相：" + when, DARK, False, 13), ("机制：" + mech, DARK, False, 13), ("兴奋性：" + exc, col, True, 13)], spacing=1.18)
add_rect(s, Inches(0.4), Inches(5.55), Inches(6.15), Inches(1.65), LIGHT, BLUE, rounded=True)
add_text(s, Inches(0.55), Inches(5.61), Inches(5.9), Inches(1.55), [("配图建议", BLUE, True, 13), (p["fig"], DARK, False, 10.5)], spacing=1.08)
add_rect(s, Inches(6.75), Inches(5.55), Inches(6.15), Inches(1.65), ORANGE_BG, ORANGE, rounded=True)
add_text(s, Inches(6.9), Inches(5.61), Inches(5.9), Inches(1.55), [("教师备注", ORANGE_TX, True, 13), (p["note"], NOTE_TX, False, 10.5)], spacing=1.08)

# P16 临床三卡
p = next(x for x in ALL if x["page"]==16); s = prs.slides.add_slide(BLANK); header(s, 16, p["title"])
cx = Inches(0.4)
for title, lines in p["blocks"]:
    add_rect(s, cx, Inches(1.0), Inches(4.15), Inches(3.4), WHITE, BLUE, rounded=True)
    add_rect(s, cx, Inches(1.0), Inches(4.15), Inches(0.5), BLUE, rounded=True)
    add_text(s, cx+Inches(0.12), Inches(1.06), Inches(3.9), Inches(0.4), [(title, WHITE, True, 14)], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, cx+Inches(0.14), Inches(1.58), Inches(3.88), Inches(2.7), [("• " + ln, DARK, False, 12) for ln in lines], spacing=1.15)
    cx += Inches(4.27)
add_rect(s, Inches(0.4), Inches(4.55), Inches(6.15), Inches(1.8), LIGHT, BLUE, rounded=True)
add_text(s, Inches(0.55), Inches(4.61), Inches(5.9), Inches(1.7), [("配图建议", BLUE, True, 13), (p["fig"], DARK, False, 10.5)], spacing=1.08)
add_rect(s, Inches(6.75), Inches(4.55), Inches(6.15), Inches(1.8), ORANGE_BG, ORANGE, rounded=True)
add_text(s, Inches(6.9), Inches(4.61), Inches(5.9), Inches(1.7), [("教师备注", ORANGE_TX, True, 13), (p["note"], NOTE_TX, False, 10.5)], spacing=1.08)

# P17 六宫格
p = next(x for x in ALL if x["page"]==17); s = prs.slides.add_slide(BLANK); header(s, 17, p["title"])
gw, gh = Inches(4.15), Inches(1.62)
for i, (t, d) in enumerate(p["items"]):
    r_, c_ = divmod(i, 3)
    px = Inches(0.4) + c_ * Inches(4.27); py = Inches(1.0) + r_ * Inches(1.74)
    add_rect(s, px, py, gw, gh, WHITE, BLUE, rounded=True)
    add_rect(s, px, py, Inches(0.5), gh, BLUE if i%2==0 else ORANGE)
    add_text(s, px+Inches(0.6), py+Inches(0.06), gw-Inches(0.7), Inches(0.35), [(f"{i+1}. {t}", BLUE if i%2==0 else ORANGE_TX, True, 13)])
    add_text(s, px+Inches(0.6), py+Inches(0.42), gw-Inches(0.72), gh-Inches(0.5), [(d, DARK, False, 10.5)], spacing=1.08)
add_rect(s, Inches(0.4), Inches(4.62), Inches(6.15), Inches(1.7), LIGHT, BLUE, rounded=True)
add_text(s, Inches(0.55), Inches(4.68), Inches(5.9), Inches(1.6), [("配图建议", BLUE, True, 12), (p["fig"], DARK, False, 10)], spacing=1.05)
add_rect(s, Inches(6.75), Inches(4.62), Inches(6.15), Inches(1.7), ORANGE_BG, ORANGE, rounded=True)
add_text(s, Inches(6.9), Inches(4.68), Inches(5.9), Inches(1.6), [("教师备注", ORANGE_TX, True, 12), (p["note"], NOTE_TX, False, 10)], spacing=1.05)

# P18 小结
p = next(x for x in ALL if x["page"]==18); s = prs.slides.add_slide(BLANK); header(s, 18, p["title"])
add_rect(s, Inches(4.8), Inches(1.0), Inches(3.7), Inches(0.7), BLUE, rounded=True)
add_text(s, Inches(4.8), Inches(1.08), Inches(3.7), Inches(0.55), [("细胞电生理 · 一页框架", WHITE, True, 18)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
y = Inches(2.0)
for grp, line in p["tree"]:
    if grp:
        col = ORANGE if grp=="临床锚点" else LIGHT; tx = ORANGE_TX if grp=="临床锚点" else BLUE
        add_rect(s, Inches(0.6), y, Inches(2.0), Inches(0.55), col, ORANGE if grp=="临床锚点" else BLUE, rounded=True)
        add_text(s, Inches(0.6), y+Inches(0.06), Inches(2.0), Inches(0.45), [(grp, tx, True, 14)], align=PP_ALIGN.CENTER, anchor=MSO_ANCHOR.MIDDLE)
    add_rect(s, Inches(2.85), y, Inches(10.05), Inches(0.55), WHITE, RGBColor(0xD8,0xE4,0xF0), rounded=True)
    add_text(s, Inches(3.0), y+Inches(0.05), Inches(9.8), Inches(0.48), [(line, DARK, False, 14)], anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.66)
add_rect(s, Inches(0.4), Inches(6.0), Inches(6.15), Inches(1.2), LIGHT, BLUE, rounded=True)
add_text(s, Inches(0.55), Inches(6.06), Inches(5.9), Inches(1.1), [("配图建议", BLUE, True, 12), (p["fig"], DARK, False, 10)], spacing=1.05)
add_rect(s, Inches(6.75), Inches(6.0), Inches(6.15), Inches(1.2), ORANGE_BG, ORANGE, rounded=True)
add_text(s, Inches(6.9), Inches(6.06), Inches(5.9), Inches(1.1), [("教师备注", ORANGE_TX, True, 12), (p["note"], NOTE_TX, False, 10)], spacing=1.05)

# P19 作业
p = next(x for x in ALL if x["page"]==19); s = prs.slides.add_slide(BLANK); header(s, 19, p["title"])
for j, group in enumerate(p["hw"]):
    t, items = group[0], group[1:]
    px = Inches(0.4) + j * Inches(6.35)
    col = RED if j==0 else BLUE
    add_rect(s, px, Inches(1.0), Inches(6.15), Inches(3.3), WHITE, col, rounded=True)
    add_rect(s, px, Inches(1.0), Inches(6.15), Inches(0.5), col, rounded=True)
    add_text(s, px+Inches(0.15), Inches(1.06), Inches(5.9), Inches(0.4), [(t, WHITE, True, 15)], anchor=MSO_ANCHOR.MIDDLE)
    add_text(s, px+Inches(0.18), Inches(1.58), Inches(5.85), Inches(2.6), [(it, DARK, False, 13) for it in items], spacing=1.3)
add_rect(s, Inches(0.4), Inches(4.5), Inches(6.15), Inches(1.8), LIGHT, BLUE, rounded=True)
add_text(s, Inches(0.55), Inches(4.56), Inches(5.9), Inches(1.7), [("配图建议", BLUE, True, 13), (p["fig"], DARK, False, 10.5)], spacing=1.08)
add_rect(s, Inches(6.75), Inches(4.5), Inches(6.15), Inches(1.8), ORANGE_BG, ORANGE, rounded=True)
add_text(s, Inches(6.9), Inches(4.56), Inches(5.9), Inches(1.7), [("教师备注", ORANGE_TX, True, 13), (p["note"], NOTE_TX, False, 10.5)], spacing=1.08)

# P20 参考
p = next(x for x in ALL if x["page"]==20); s = prs.slides.add_slide(BLANK); header(s, 20, p["title"])
y = Inches(1.1)
for ref in p["refs"]:
    add_rect(s, Inches(0.6), y, Inches(12.1), Inches(0.72), LIGHT, RGBColor(0xE2,0xEC,0xF6), rounded=True)
    add_text(s, Inches(0.8), y+Inches(0.07), Inches(11.7), Inches(0.6), [(ref, DARK, False, 14)], anchor=MSO_ANCHOR.MIDDLE)
    y += Inches(0.82)
add_rect(s, Inches(0.4), Inches(6.1), Inches(6.15), Inches(1.15), LIGHT, BLUE, rounded=True)
add_text(s, Inches(0.55), Inches(6.15), Inches(5.9), Inches(1.05), [("配图建议", BLUE, True, 12), (p["fig"], DARK, False, 10)], spacing=1.05)
add_rect(s, Inches(6.75), Inches(6.1), Inches(6.15), Inches(1.15), ORANGE_BG, ORANGE, rounded=True)
add_text(s, Inches(6.9), Inches(6.15), Inches(5.9), Inches(1.05), [("教师备注", ORANGE_TX, True, 12), (p["note"], NOTE_TX, False, 10)], spacing=1.05)

out = os.path.join(D, "生理学-细胞的电生理现象-教学PPT.pptx")
prs.save(out)
print("OK", out, "slides:", len(prs.slides._sldIdLst), "bytes:", os.path.getsize(out))