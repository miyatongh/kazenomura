#!/usr/bin/env python3
"""
2026/2/16 Phase 0 説明会（思考公開型）スライド生成
- 30枚構成（説明40分 + 質疑20分）
"""

from pathlib import Path
from pptx import Presentation
from pptx.util import Inches, Pt, Emu
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

SLIDE_WIDTH = Inches(13.333)
SLIDE_HEIGHT = Inches(7.5)
TOTAL = 30

C_NAVY = RGBColor(0x1B, 0x2A, 0x4A)
C_DARK = RGBColor(0x2C, 0x3E, 0x50)
C_BLUE = RGBColor(0x2E, 0x86, 0xC1)
C_TEAL = RGBColor(0x17, 0xA5, 0x89)
C_ORANGE = RGBColor(0xE6, 0x7E, 0x22)
C_WHITE = RGBColor(0xFF, 0xFF, 0xFF)
C_GRAY = RGBColor(0x7F, 0x8C, 0x8D)

prs = Presentation()
prs.slide_width = SLIDE_WIDTH
prs.slide_height = SLIDE_HEIGHT


def add_bg(slide, color):
    shape = slide.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(0), Emu(0), SLIDE_WIDTH, SLIDE_HEIGHT)
    shape.fill.solid()
    shape.fill.fore_color.rgb = color
    shape.line.fill.background()


def add_text(slide, left, top, width, height, text, size=18, bold=False, color=C_DARK, align=PP_ALIGN.LEFT):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    p = tf.paragraphs[0]
    p.text = text
    p.font.size = Pt(size)
    p.font.bold = bold
    p.font.color.rgb = color
    p.font.name = "Yu Gothic"
    p.alignment = align


def add_bullets(slide, left, top, width, height, bullets):
    box = slide.shapes.add_textbox(left, top, width, height)
    tf = box.text_frame
    tf.word_wrap = True
    for i, b in enumerate(bullets):
        p = tf.paragraphs[0] if i == 0 else tf.add_paragraph()
        if isinstance(b, dict):
            p.text = b["text"]
            p.font.size = Pt(b.get("size", 20))
            p.font.bold = b.get("bold", False)
            p.font.color.rgb = b.get("color", C_DARK)
        else:
            p.text = b
            p.font.size = Pt(20)
            p.font.bold = False
            p.font.color.rgb = C_DARK
        p.font.name = "Yu Gothic"
        p.space_after = Pt(8)


def add_page(slide, num):
    add_text(slide, Inches(11.8), Inches(6.9), Inches(1.2), Inches(0.4), f"{num} / {TOTAL}", size=10, color=C_GRAY, align=PP_ALIGN.RIGHT)


def title_slide():
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(s, C_NAVY)
    add_text(s, Inches(1), Inches(2.0), Inches(11.3), Inches(1.2), "Phase 0『要求設計』コンサルティング説明会", size=34, bold=True, color=C_WHITE)
    add_text(s, Inches(1), Inches(3.2), Inches(11.3), Inches(0.8), "思考公開型：結果を生む考え方を共有する", size=22, color=RGBColor(0xAE, 0xBF, 0xD5))
    add_text(s, Inches(1), Inches(5.4), Inches(5), Inches(0.5), "2026年2月16日", size=16, color=RGBColor(0xAE, 0xBF, 0xD5))
    add_text(s, Inches(7), Inches(5.4), Inches(5.3), Inches(0.5), "株式会社PreSoft / eMu 上田昌夫", size=16, color=RGBColor(0xAE, 0xBF, 0xD5), align=PP_ALIGN.RIGHT)


def section_slide(number, title, subtitle):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    add_bg(s, C_NAVY)
    add_text(s, Inches(1), Inches(1.5), Inches(2), Inches(2.2), str(number), size=92, bold=True, color=C_BLUE)
    add_text(s, Inches(1), Inches(3.5), Inches(11), Inches(1.0), title, size=34, bold=True, color=C_WHITE)
    add_text(s, Inches(1), Inches(4.6), Inches(11), Inches(0.8), subtitle, size=20, color=RGBColor(0xAE, 0xBF, 0xD5))


def content_slide(num, title, bullets, note=""):
    s = prs.slides.add_slide(prs.slide_layouts[6])
    head = s.shapes.add_shape(MSO_SHAPE.RECTANGLE, Emu(0), Emu(0), SLIDE_WIDTH, Inches(1.2))
    head.fill.solid()
    head.fill.fore_color.rgb = C_NAVY
    head.line.fill.background()
    add_text(s, Inches(0.8), Inches(0.15), Inches(11.5), Inches(0.9), title, size=28, bold=True, color=C_WHITE)
    add_bullets(s, Inches(1.0), Inches(1.6), Inches(11.3), Inches(4.9), bullets)
    if note:
        add_text(s, Inches(1.0), Inches(6.5), Inches(11.3), Inches(0.5), note, size=14, color=C_GRAY)
    add_page(s, num)


def build():
    title_slide()

    slides = [
        (2, "本日の位置づけ", [
            "12月・1月で『課題』と『方向性』は共有済みです。",
            {"text": "本日は、結果ではなく『判断の考え方』を共有します。", "bold": True, "color": C_ORANGE},
            "狙いは、幹部間で再利用可能な意思決定の物差しをそろえることです。",
        ], ""),
        (3, "先に結論", [
            {"text": "今回の要点は『何をしたか』より『どう考えるか』です。", "bold": True, "color": C_ORANGE},
            "思考が違えば、同じ課題でも結論は変わります。",
            "思考が共有されると、結果の良し悪しに関わらず次の判断に活かせます。",
        ], ""),
        (4, "一般的な説明との違い", [
            "一般的な説明：What（何をしたか）/ Result（結果）中心",
            {"text": "今回の説明：Why（なぜ）/ How（どう判断するか）中心", "bold": True, "color": C_BLUE},
            "『思考→活動→成果』の順でご説明します。",
        ], ""),
        (5, "なぜ思考公開が必要か", [
            "結果だけの共有では、次回の再現性が低くなります。",
            "思考手順を共有すると、組織として判断を積み上げられます。",
            {"text": "今回のコンサルティングは、思考の共有自体が成果です。", "bold": True, "color": C_ORANGE},
        ], ""),
        (6, "本日の4テーマ", [
            "① 我々の基本的考え方／方法論",
            "② 何のために何をするか（Phase 0活動）",
            "③ 何を作って何を見せるか（成果物）",
            "④ 幹部の皆さまにお願いしたいこと",
        ], "説明40分 + 質疑20分"),

        (8, "思考の起点：課題は個人でなく構造", [
            "多重入力・属人化・不可視化は、個人能力の問題ではありません。",
            {"text": "構造の問題として捉えることが、改革の出発点です。", "bold": True, "color": C_ORANGE},
            "したがって解決策は、個人努力でなく業務設計に置きます。",
        ], ""),
        (9, "三原理は『目標』でなく『前提』", [
            "PDマネジメント（計画と実績）",
            "知識と作業の分離",
            "リアルタイム一個流し",
            {"text": "これらは追加要件ではなく、本来あるべき業務構造です。", "bold": True, "color": C_BLUE},
        ], ""),
        (10, "改革の定義", [
            "新しいことを足す改革ではありません。",
            {"text": "本来つながっている業務を、つなぎ直す改革です。", "bold": True, "color": C_ORANGE},
            "『導入』より先に『構造を明らかにする』を置きます。",
        ], ""),
        (11, "判断順序（最重要）", [
            {"text": "仕事（業務構造）→ 機能 → システム", "bold": True, "color": C_BLUE},
            "この順序を逆にすると、部分最適と手作業が増えます。",
            "Phase 0は、この順序を守るための設計フェーズです。",
        ], ""),
        (12, "4視点で全体を捉える", [
            "業務の流れ（誰が何をいつ）",
            "情報の流れ（どこで生まれどこへ届くか）",
            "道具の配置（機能をどこに置くか）",
            "基盤（運用と拡張を支える土台）",
        ], ""),
        (13, "3階層で断絶を見つける", [
            "サービス提供の層",
            "経営資源を管理する層",
            "現場で実行する層",
            {"text": "層間の断絶が、多重入力と遅延の根本原因です。", "bold": True, "color": C_ORANGE},
        ], ""),
        (14, "翻訳ルール：聴く→構造化→要件化", [
            "個別意見をそのまま要件化しません。",
            "共通パターンを抽出して、構造課題に変換します。",
            "その上で業務要件・運用要件へ落とし込みます。",
        ], ""),
        (15, "この方法論で防げること", [
            "場当たり導入の固定化",
            "月末集中と遅延の常態化",
            "ブラックボックス化した運用",
            {"text": "判断の再現性を確保し、組織学習を可能にします。", "bold": True, "color": C_BLUE},
        ], ""),

        (17, "Phase 0の目的（再定義）", [
            "目的は『資料を作ること』ではありません。",
            {"text": "目的は『業務構造改革を実行できる設計』を確立することです。", "bold": True, "color": C_ORANGE},
            "何から着手し、どう展開するかを経営判断可能にします。",
        ], ""),
        (18, "3か月の活動設計", [
            "ヒアリング調査：25〜30回（管理本部 + 現場）",
            "隔週プロジェクト会議：6回",
            "分析・設計：全期間で反復実施",
        ], ""),
        (19, "なぜこの活動設計か", [
            "拠点を網羅して『偏った設計』を避けるため",
            "つなぎ目の断絶を実地で特定するため",
            "共通パターンを抽出して設計精度を上げるため",
        ], ""),
        (20, "進め方の特徴", [
            "密室で作らず、途中共有・途中修正を前提に進めます。",
            "隔週会議で判断論点を明示し、遅延を防ぎます。",
            {"text": "『見えないコンサル』にしないことを重視します。", "bold": True, "color": C_BLUE},
        ], ""),
        (21, "Phase 0でやらないこと", [
            "システム導入・個別実装は行いません。",
            "ベンダー選定の確定もこの段階では行いません。",
            {"text": "先に設計図を確定し、次フェーズで実装判断します。", "bold": True, "color": C_ORANGE},
        ], ""),

        (23, "成果物の全体像", [
            "① マスタープラン",
            "② 要件定義書",
            "③ マスターデータ構造設計",
            "3つは『読む資料』でなく『決める資料』です。",
        ], ""),
        (24, "成果物① マスタープラン", [
            "全体像・優先順位・段階展開を示します。",
            "どこから着手し、どこまでを次フェーズに送るかを判断できます。",
            "概算観点を持ち、投資対効果の議論が可能になります。",
        ], ""),
        (25, "成果物② 要件定義書", [
            "業務要件 / 運用要件 / 機能要件 / システム要件の4階層",
            "要望を仕様に翻訳し、ベンダー比較可能性を確保します。",
            {"text": "『言った・言わない』を防ぐ共通基盤になります。", "bold": True, "color": C_BLUE},
        ], ""),
        (26, "成果物③ マスターデータ構造設計", [
            "One Fact, One Place を実現するデータ設計",
            "多重入力・不整合・転記ミスを削減",
            "将来の連携拡張の土台を作ります。",
        ], ""),

        (28, "4つのお願い", [
            "ヒアリング対象者の選定と日程調整",
            "既存資料（規程・帳票・運用資料）の提供",
            "隔週プロジェクト会議への参加",
            "責任者・窓口・意思決定ルートの明確化",
        ], ""),
        (29, "なぜ幹部関与が必須か", [
            "前提・優先順位・範囲の判断は幹部の役割です。",
            "ここが曖昧だと、現場は従来運用へ戻ります。",
            {"text": "思考の合意が、改革実行性を決めます。", "bold": True, "color": C_ORANGE},
        ], ""),
        (30, "まとめ：本日の合意事項", [
            "① 結果より先に思考を共有する",
            "② Phase 0は実行設計を作るフェーズである",
            "③ 成果物は経営判断の道具である",
            "④ 幹部は意思決定プロセスのオーナーである",
        ], "ご確認後、キックオフに向けた体制確定へ進みます"),
    ]

    section_marks = {
        7: ("①", "基本的な考え方／方法論", "我々の思考OSを共有する"),
        16: ("②", "何のために何をするか", "思考を実務に落とす"),
        22: ("③", "何を作って何を見せるか", "思考の出力としての成果物"),
        27: ("④", "幹部にお願いすること", "思考を組織で機能させる条件"),
    }

    for n in range(2, TOTAL + 1):
        if n in section_marks:
            sec = section_marks[n]
            section_slide(*sec)
            continue
        found = next((x for x in slides if x[0] == n), None)
        if found:
            _, title, bullets, note = found
            content_slide(n, title, bullets, note)
        else:
            content_slide(n, "（予備スライド）", [""], "")


if __name__ == "__main__":
    build()
    out = Path(__file__).resolve().parents[1] / "03_成果物" / "20260216_Phase0説明会_本編_v02_思考公開型.pptx"
    prs.save(str(out))
    print(f"Created: {out}")
