# -*- coding: utf-8 -*-
"""make_slides.py — 產生期末報告用的中文簡報 (12 張, 16:9)。

敘事主軸對齊 paper.tex 的核心發現:「在安全攸關的 RL 中,episodic return 是會騙人的指標」。
圖表直接嵌入 logs/ 內已產生的學習曲線。

用法 (從 repo root):
    python src/make_slides.py --out submit/N26140511_slides.pptx
"""
import argparse
import os

from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR

# ---- 主題色 ----
NAVY = RGBColor(0x1F, 0x2D, 0x49)      # 標題深藍
ACCENT = RGBColor(0xC6, 0x3A, 0x2E)    # 強調紅 (對應 CRASHED)
GREEN = RGBColor(0x1E, 0x7A, 0x3C)     # DQN 安全綠
GREY = RGBColor(0x55, 0x5B, 0x66)
WHITE = RGBColor(0xFF, 0xFF, 0xFF)
LIGHT = RGBColor(0xF2, 0xF4, 0xF7)

CN = "Microsoft JhengHei"   # 微軟正黑體
CN_L = "Microsoft JhengHei Light"

prs = Presentation()
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)
SW, SH = prs.slide_width, prs.slide_height
BLANK = prs.slide_layouts[6]


def _set_font(run, size, bold=False, color=NAVY, font=CN):
    run.font.size = Pt(size)
    run.font.bold = bold
    run.font.color.rgb = color
    run.font.name = font
    # 同時設定東亞字型,避免中文 fallback
    rPr = run._r.get_or_add_rPr()
    import copy
    ea = rPr.find('{http://schemas.openxmlformats.org/drawingml/2006/main}ea')
    if ea is None:
        from pptx.oxml.ns import qn
        ea = rPr.makeelement(qn('a:ea'), {})
        rPr.append(ea)
    ea.set('typeface', font)


def add_band(slide, color=NAVY, height=Inches(0.18), top=None):
    top = SH - height if top is None else top
    box = slide.shapes.add_shape(1, 0, top, SW, height)
    box.fill.solid(); box.fill.fore_color.rgb = color
    box.line.fill.background()
    box.shadow.inherit = False
    return box


def title_slide():
    s = prs.slides.add_slide(BLANK)
    # 上方深藍塊
    band = s.shapes.add_shape(1, 0, 0, SW, Inches(4.7))
    band.fill.solid(); band.fill.fore_color.rgb = NAVY
    band.line.fill.background(); band.shadow.inherit = False

    tb = s.shapes.add_textbox(Inches(0.9), Inches(1.0), Inches(11.5), Inches(3.4))
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = "當回合獎勵會騙人"
    _set_font(r, 44, bold=True, color=WHITE)
    p2 = tf.add_paragraph()
    r = p2.add_run(); r.text = "PPO 與 DQN 在 highway-env 上的行為消融研究"
    _set_font(r, 24, bold=False, color=RGBColor(0xC9, 0xD3, 0xE6))
    p3 = tf.add_paragraph(); p3.space_before = Pt(18)
    r = p3.add_run()
    r.text = "Return parity ≤5.7%,碰撞率卻是 11.3% vs 60–80%"
    _set_font(r, 18, bold=False, color=ACCENT, font=CN)

    # 下方資訊
    info = s.shapes.add_textbox(Inches(0.9), Inches(5.1), Inches(11.5), Inches(1.8))
    tf = info.text_frame; tf.word_wrap = True
    lines = [
        ("阮紹銘 (Shao-Ming Ruan) ・ 學號 N26140511", 18, NAVY, True),
        ("國立成功大學 電機工程研究所 ・ Group F", 15, GREY, False),
        ("強化學習 期末報告 ・ 2026/06", 14, GREY, False),
    ]
    for i, (txt, sz, col, bd) in enumerate(lines):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        r = p.add_run(); r.text = txt
        _set_font(r, sz, bold=bd, color=col)
    add_band(s, ACCENT)
    return s


def content_slide(title, kicker=None):
    s = prs.slides.add_slide(BLANK)
    # 標題列
    tb = s.shapes.add_textbox(Inches(0.7), Inches(0.4), Inches(11.9), Inches(1.0))
    tf = tb.text_frame; tf.word_wrap = True
    p = tf.paragraphs[0]
    r = p.add_run(); r.text = title
    _set_font(r, 30, bold=True, color=NAVY)
    if kicker:
        p2 = tf.add_paragraph()
        r = p2.add_run(); r.text = kicker
        _set_font(r, 14, bold=False, color=ACCENT)
    # 標題底線
    ln = s.shapes.add_shape(1, Inches(0.7), Inches(1.5), Inches(11.9), Pt(2.5))
    ln.fill.solid(); ln.fill.fore_color.rgb = NAVY
    ln.line.fill.background(); ln.shadow.inherit = False
    add_band(s)
    return s


def bullets(slide, items, left=Inches(0.8), top=Inches(1.8),
            width=Inches(11.7), height=Inches(5.0), size=18):
    """items: list of (text, level, color, bold)"""
    tb = slide.shapes.add_textbox(left, top, width, height)
    tf = tb.text_frame; tf.word_wrap = True
    for i, item in enumerate(items):
        txt, lvl, col, bd = (item + (None, None))[:4]
        col = col or NAVY
        bd = bool(bd)
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        p.level = lvl
        p.space_after = Pt(8)
        bullet = ("● " if lvl == 0 else "– ")
        r = p.add_run(); r.text = bullet + txt
        _set_font(r, size - lvl * 2, bold=bd, color=col)
    return tb


def add_image(slide, path, left, top, width=None, height=None):
    if not os.path.exists(path):
        return None
    return slide.shapes.add_picture(path, left, top, width=width, height=height)


def metric_card(slide, left, top, w, h, label, value, color):
    box = slide.shapes.add_shape(1, left, top, w, h)
    box.fill.solid(); box.fill.fore_color.rgb = color
    box.line.fill.background(); box.shadow.inherit = False
    tf = box.text_frame; tf.word_wrap = True
    tf.vertical_anchor = MSO_ANCHOR.MIDDLE
    p = tf.paragraphs[0]; p.alignment = PP_ALIGN.CENTER
    r = p.add_run(); r.text = value
    _set_font(r, 30, bold=True, color=WHITE)
    p2 = tf.add_paragraph(); p2.alignment = PP_ALIGN.CENTER
    r = p2.add_run(); r.text = label
    _set_font(r, 13, bold=False, color=WHITE)
    return box


def simple_table(slide, data, left, top, width, height,
                 header_color=NAVY, font_size=14, highlight_row=None):
    rows, cols = len(data), len(data[0])
    gtbl = slide.shapes.add_table(rows, cols, left, top, width, height)
    tbl = gtbl.table
    for r_i, row in enumerate(data):
        for c_i, val in enumerate(row):
            cell = tbl.cell(r_i, c_i)
            cell.margin_top = Pt(3); cell.margin_bottom = Pt(3)
            cell.margin_left = Pt(6); cell.margin_right = Pt(6)
            cell.vertical_anchor = MSO_ANCHOR.MIDDLE
            tf = cell.text_frame; tf.word_wrap = True
            p = tf.paragraphs[0]
            p.alignment = PP_ALIGN.LEFT if c_i == 0 else PP_ALIGN.CENTER
            run = p.add_run(); run.text = str(val)
            if r_i == 0:
                _set_font(run, font_size, bold=True, color=WHITE)
                cell.fill.solid(); cell.fill.fore_color.rgb = header_color
            else:
                hl = (highlight_row is not None and r_i == highlight_row)
                _set_font(run, font_size, bold=hl,
                          color=(GREEN if hl else NAVY))
                cell.fill.solid()
                cell.fill.fore_color.rgb = (RGBColor(0xE4, 0xF2, 0xE8) if hl
                                            else (LIGHT if r_i % 2 else WHITE))
    return tbl


FIG = "logs"


def build():
    # 1 ── 標題
    title_slide()

    # 2 ── 研究動機
    s = content_slide("研究動機:我們評估 RL 的方式可靠嗎?",
                      "Motivation")
    bullets(s, [
        ("自駕決策(何時變換車道、加速、禮讓)天然是一個序列決策問題 → 適合深度強化學習。", 0),
        ("highway-env 是此類問題的標準輕量 benchmark:小型運動學狀態、離散 meta-action、可調獎勵。", 0),
        ("實務上幾乎所有論文都只報告「平均回合獎勵 (episodic return)」作為成效指標。", 0),
        ("但在安全攸關的控制裡 — return 高,就代表車開得安全嗎?", 0, ACCENT, True),
        ("本研究在「參數對齊」的公平條件下,量化 return 與實際駕駛行為之間的落差。", 0, NAVY, True),
    ], size=19)

    # 3 ── 問題與貢獻
    s = content_slide("研究問題與核心貢獻", "Research Questions & Contributions")
    bullets(s, [
        ("兩個少被嚴格檢驗的設計選擇:", 0, NAVY, True),
        ("特徵抽取器:Self-Attention(集合、置換等變)vs 純 MLP — 同參數預算下誰較好?", 1),
        ("演算法家族:On-policy 的 PPO vs Off-policy 的 Double DQN — 樣本效率與行為?", 1),
        ("核心貢獻(實驗發現,而非方法創新):", 0, ACCENT, True),
        ("三種配置的 return 差距 ≤5.7%,但部署碰撞率相差一個數量級(11.3% vs 60–80%)。", 1),
        ("此結論泛化到第二個環境 merge-v0(return 差 ≤1.4%,碰撞率 0% vs 100%)。", 1),
        ("主張:行為層級評估是 return 曲線的必要補充,而非選配。", 1, NAVY, True),
    ], size=18)

    # 4 ── 環境與設定
    s = content_slide("環境與實驗設定", "highway-env · DiscreteMetaAction")
    bullets(s, [
        ("觀測:Kinematics,5 台車 × 5 特徵(presence, x, y, vx, vy),ego 相對且正規化。", 0),
        ("動作:DiscreteMetaAction 五個高階操作(左/右變道、IDLE、加速、減速),5 Hz 決策。", 0),
        ("回合長度:duration=80 模擬秒 = 最多 400 個決策步(非 80 步)。", 0),
        ("三種獎勵風格(僅差在獎勵係數):base / conservative / aggressive。", 0),
        ("每組 5 個 seed、各訓練 100,000 步;評估跑 30 局 deterministic。", 0),
    ], top=Inches(1.7), height=Inches(2.6), size=18)
    simple_table(s, [
        ["風格", "碰撞罰", "高速獎勵", "變道", "速度區間 (m/s)"],
        ["base", "−1.0", "0.4", "0.0", "[20, 30]"],
        ["conservative", "−2.0", "0.4", "−0.5", "[10, 20]"],
        ["aggressive", "−4.0", "2.0", "+0.2", "[30, 40]"],
    ], Inches(2.4), Inches(4.5), Inches(8.5), Inches(1.9), font_size=14)

    # 5 ── 方法
    s = content_slide("方法:從零實作 PPO 與 Double DQN", "Methods")
    bullets(s, [
        ("PPO(手刻):GAE(γ=0.99, λ=0.95)+ clipped surrogate(ε=0.2)+ 每 minibatch 優勢正規化 + 線性 LR 衰減。", 0),
        ("關鍵實作:GAE 的 nextnonterminal 終止遮罩必須正確,否則優勢會跨回合洩漏 → 曲線發散。", 1, ACCENT, False),
        ("Double DQN:replay buffer(5 萬)+ target network(每 500 步硬更新)+ ε-greedy + double-Q 目標。", 0),
        ("以 Stable-Baselines3 PPO 作為正確性交叉驗證(同超參、seed 0)。", 0),
        ("三個網路皆吃 5×5 觀測;Attention 與 MLP 參數對齊(消融才公平):", 0, NAVY, True),
    ], top=Inches(1.7), height=Inches(3.0), size=17)
    simple_table(s, [
        ["網路", "參數量", "角色"],
        ["AttentionActorCritic", "66,822", "PPO actor–critic(注意力)"],
        ["MlpActorCritic", "72,262", "PPO actor–critic(MLP)"],
        ["MlpQNetwork", "20,485", "DQN 價值網路(刻意較小)"],
    ], Inches(2.6), Inches(4.9), Inches(8.1), Inches(1.7), font_size=14)

    # 6 ── 結果一:return 打平
    s = content_slide("結果一:Return 幾乎打平", "highway-v0 · aggressive · n=5 seeds")
    simple_table(s, [
        ["配置", "Final reward", "Peak", "到 avg≥150 步數", "跨 seed σ"],
        ["PPO + Attention", "235.7 ± 16.7", "266.7", "≈10,240", "16.7"],
        ["PPO + MLP", "223.0 ± 19.6", "266.7", "≈9,216", "19.6"],
        ["DQN + MLP", "234.1 ± 7.7", "266.7", "≈7,721", "7.7(最低)"],
    ], Inches(0.7), Inches(1.8), Inches(6.5), Inches(2.2),
        font_size=13, highlight_row=3)
    bullets(s, [
        ("三者差距 ≤5.7%,都逼近理論上限 ~266.7。", 0, NAVY, True),
        ("DQN 跨 seed 變異最低,且最早到 avg≥150(比 Attn 早 ~25%)。", 0),
        ("歸因:replay 平滑早期高變異梯度。", 0),
        ("⇒ 架構/演算法幾乎不決定「能爬多高」。", 0, ACCENT, True),
    ], left=Inches(0.7), top=Inches(4.3), width=Inches(6.5),
        height=Inches(2.6), size=15)
    add_image(s, os.path.join(FIG, "comparison_aggressive.png"),
              Inches(7.5), Inches(1.9), width=Inches(5.4))

    # 7 ── 結果二:行為天差地遠(punchline)
    s = content_slide("結果二:行為卻天差地遠", "← 整份報告的核心")
    metric_card(s, Inches(0.8), Inches(1.9), Inches(3.6), Inches(1.7),
                "PPO + Attention 碰撞率", "60.7%", ACCENT)
    metric_card(s, Inches(4.7), Inches(1.9), Inches(3.6), Inches(1.7),
                "PPO + MLP 碰撞率", "80.0%", ACCENT)
    metric_card(s, Inches(8.6), Inches(1.9), Inches(3.6), Inches(1.7),
                "DQN + MLP 碰撞率", "11.3%", GREEN)
    bullets(s, [
        ("訓練 return 幾乎相同,部署碰撞率卻差一個數量級。", 0, NAVY, True),
        ("PPO 呈雙峰:部分 seed 完全保守(0–3.3% 撞),部分災難性(100% 撞)。", 0),
        ("DQN 明顯較一致(11.3%),存活步數 368.5(PPO 僅 152–200)。", 0),
        ("為何 return 看不出來?aggressive 每步 +2.0 高速獎勵 → 中後期才撞的快車,累積 return", 0),
        ("與全程存活的慢車相當;單次 −4.0 終止罰相對微小。(早撞仍然昂貴。)", 1),
        ("⚠️ DQN 低碰撞伴隨極頻繁變道 → 不排除「靠甩尾求生」的退化策略(open caveat)。", 0, GREY, False),
    ], top=Inches(3.9), height=Inches(2.9), size=15)

    # 8 ── 結果三:merge-v0 泛化
    s = content_slide("結果三:泛化到第二個環境 merge-v0", "Generalization")
    bullets(s, [
        ("merge-v0 動態刻意與 highway 最不像:從短匝道切入、匝道末端有障礙物。", 0),
        ("若結論在此仍成立,「只在 highway 才對」的反駁就難以成立。", 0),
    ], top=Inches(1.7), height=Inches(1.2), size=17)
    simple_table(s, [
        ["配置", "訓練 reward", "碰撞率", "存活步數", "平均速度"],
        ["PPO + Attention", "58.3 ± 0.2", "100%", "34.0", "29.9 m/s"],
        ["PPO + MLP", "59.1 ± 0.3", "100%", "35.0", "29.6 m/s"],
        ["DQN + MLP", "59.0 ± 0.3", "0.0%", "77.7", "22.1 m/s"],
    ], Inches(1.4), Inches(3.0), Inches(10.5), Inches(2.0),
        font_size=14, highlight_row=3)
    bullets(s, [
        ("訓練 reward 更緊(差 ≤1.4%,σ≤0.3),行為卻是 全有/全無:DQN 0% vs PPO 100%。", 0, NAVY, True),
        ("且無雙峰 — PPO 在每個 seed 都收斂到最高速撞車策略。⇒ 強化核心主張。", 0, ACCENT, True),
    ], top=Inches(5.3), height=Inches(1.6), size=16)

    # 9 ── 獎勵塑形
    s = content_slide("獎勵塑形:主控桿,但有變異性警示", "Reward-shaping study")
    bullets(s, [
        ("同一架構(PPO+Attention)、不同獎勵風格 → 駕駛風格天差地遠。", 0),
        ("base 訓練最安全(0% 撞、400 步、0 變道);aggressive 最快但最常撞。", 0),
        ("conservative「想要」更安全,卻雙峰:seed 0,1 完美(0%),seed 2 全撞(100%)。", 0, ACCENT, False),
        ("⇒ 獎勵設計約束學習「方向」,但不保證所有 seed 收斂到目標 basin。", 0, NAVY, True),
        ("實務啟示:獎勵工程是主要槓桿,但必須搭配「多 seed 的行為層級驗證」。", 0, NAVY, True),
    ], top=Inches(1.7), height=Inches(2.6), size=17)
    add_image(s, os.path.join(FIG, "comparison_styles.png"),
              Inches(3.4), Inches(4.4), width=Inches(6.5))

    # 10 ── 影片佐證
    s = content_slide("影片佐證:相同場景、相反行為", "Side-by-side demo")
    bullets(s, [
        ("兩支並排影片(上=DQN 存活、下=PPO 撞車,紅字 CRASHED),相同 reset seed → 交通完全相同,只有策略不同。", 0),
        ("demo_highway_dqn_vs_ppo.mp4 — highway-v0:DQN 存活 vs PPO+Attention 衝撞(對應 11.3% vs 60.7%)。", 0),
        ("demo_merge_dqn_vs_ppo.mp4 — merge-v0:DQN 完成合流 vs PPO 100% 撞車。", 0),
        ("誠實揭露:並排渲染會交錯兩個環境實例 → 擾動 NPC 隨機性,故畫面中的「變道次數」與表格的聚合值不必然一致;影片僅用於呈現碰撞對比。", 0, GREY, False),
    ], size=18)

    # 11 ── 討論與限制
    s = content_slide("討論與限制", "Discussion & Limitations")
    bullets(s, [
        ("為何 DQN 訓練最穩?replay buffer 維持多樣轉移,避免 Q 網路過擬合短期災難序列。", 0),
        ("行為落差是「學到的策略」本身的差異 — 行為指標對測試期獎勵不變(deterministic + 風格僅差係數),非測試分布假象。", 0, NAVY, True),
        ("限制:", 0, ACCENT, True),
        ("n=5(merge 僅 n=3),不足以做嚴格顯著性檢定;結論皆限定於所測條件。", 1),
        ("DQN 刻意較小、非容量對齊 → PPO-vs-DQN 是「演算法家族」比較。", 1),
        ("行為比較僅用單一(aggressive)訓練獎勵;swerve-to-survive 仍待軌跡分析釐清。", 1),
    ], size=17)

    # 12 ── 結論
    s = content_slide("結論", "Conclusion")
    bullets(s, [
        ("① Return 在架構與演算法間幾近相同(≤5.7%,皆逼近 ~266.7)。", 0, NAVY, True),
        ("② Return 打平卻掩蓋一個數量級的行為落差:碰撞 11.3% vs 60–80%,且泛化到 merge-v0(0% vs 100%)。", 0, NAVY, True),
        ("③ 獎勵塑形是主控桿,但同一獎勵下不同 seed 仍可能收斂到相反策略。", 0, NAVY, True),
        ("核心主張:在安全攸關的 RL,episodic return 是不足的指標 — 必須搭配多 seed 的行為層級評估。", 0, ACCENT, True),
        ("未來工作:在多種獎勵下分別訓練各演算法、擴展到 roundabout / intersection、測試 prioritized-replay DQN。", 0, GREY, False),
        ("所有程式碼與訓練模型皆已釋出,可完整重現。", 0, GREY, False),
    ], size=17)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--out", default="submit/N26140511_slides.pptx")
    args = ap.parse_args()
    build()
    os.makedirs(os.path.dirname(args.out) or ".", exist_ok=True)
    prs.save(args.out)
    print(f"Saved {len(prs.slides._sldIdLst)} slides -> {args.out}")
