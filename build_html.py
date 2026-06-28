#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""ganbari-card-mumonjuku.html（1ヶ月仕様・絵文字なし・夢門塾ブランド）を生成。
pptx版(build_pptx.py)と同じデザイン。ロゴはbase64で埋め込み自己完結。"""
import os
HERE=os.path.dirname(os.path.abspath(__file__))
import base64
logo=base64.b64encode(open(os.path.join(HERE,'mumonjuku-logo.png'),'rb').read()).decode()
LOGO=f"data:image/png;base64,{logo}"
SLOTS=20
COLORS=["sky","pink","pur","grn","yel","org"]
GAUGES=[("おだやかな こえ","落ち着いた 声で 話す"),
        ("やさしい 言葉","ふわふわ 言葉を 使う"),
        ("手・足は おだやかに","たたかない・けらない"),
        ("すわって とりくむ","自分の 場所で おちついて"),
        ("きらり ポイント","とくい・成長した こと"),
        ("今月の チャレンジ","自分で きめた こと")]

def track():
    out=[]
    for k in range(1,SLOTS+1):
        if k in (5,10,15): out.append('<span class="slot m"><i class="st"></i></span>')
        elif k==SLOTS:     out.append('<span class="slot goal"><i class="st"></i></span>')
        else:              out.append(f'<span class="slot">{k}</span>')
    return "".join(out)

gauges_html=""
for i,(n1,n2) in enumerate(GAUGES):
    c=COLORS[i]
    gauges_html+=f'''
      <div class="gauge c-{c}">
        <div class="badge b-{c}">{i+1}</div>
        <div class="gname"><div class="n1">{n1}</div><div class="n2">{n2}</div></div>
        <div class="gtrack">{track()}</div>
      </div>'''

HTML=f'''<!DOCTYPE html>
<html lang="ja"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>わたしの がんばりカード（夢門塾・1ヶ月）</title>
<style>
*{{margin:0;padding:0;box-sizing:border-box;}}
:root{{--sky:#4EC5E0;--sky-l:#7DD8EC;--sky-d:#2BA8C9;--navy:#1A3A6B;--green:#22C55E;--green-d:#15803d;
--orange:#FF8C42;--pink:#E8197D;--pink-l:#F687C0;--purple:#8B5CF6;--yellow:#FFC93C;--line:#c3d0dd;--soft:#f3f9fc;}}
html,body{{font-family:'Hiragino Maru Gothic ProN','Hiragino Kaku Gothic ProN','Hiragino Sans',sans-serif;
color:var(--navy);background:#dceef5;-webkit-print-color-adjust:exact;print-color-adjust:exact;}}
.screenbar{{max-width:1180px;margin:18px auto 8px;padding:0 16px;display:flex;gap:12px;align-items:center;flex-wrap:wrap;}}
.screenbar h1{{font-size:18px;}}.screenbar p{{font-size:13px;color:#577;flex:1 1 240px;line-height:1.5;}}
.pbtn{{background:var(--pink);color:#fff;border:none;border-radius:10px;padding:11px 22px;font-size:15px;
font-weight:800;cursor:pointer;font-family:inherit;}}
@page{{size:A4 landscape;margin:7mm;}}
.sheet{{width:283mm;min-height:196mm;margin:10px auto 40px;background:#fff;border:3px solid var(--sky);
border-radius:14px;box-shadow:0 8px 26px rgba(26,58,107,.15);padding:0 0 4mm;overflow:hidden;}}
.brandhead{{position:relative;background:linear-gradient(135deg,var(--sky),var(--sky-l));padding:4mm 6mm;
display:flex;align-items:center;gap:5mm;overflow:hidden;}}
.cloud{{position:absolute;background:rgba(255,255,255,.65);border-radius:50%;pointer-events:none;}}
.cloud.c1{{width:30mm;height:30mm;top:-16mm;left:34mm;}}.cloud.c2{{width:22mm;height:22mm;bottom:-14mm;left:120mm;}}
.cloud.c3{{width:26mm;height:26mm;top:-15mm;right:44mm;opacity:.6;}}
.logo{{flex:0 0 auto;background:#fff;border-radius:16px;padding:2mm;box-shadow:0 4px 12px rgba(26,58,107,.22);position:relative;z-index:1;}}
.logo img{{display:block;width:25mm;height:25mm;border-radius:10px;}}
.htitle{{flex:1 1 auto;position:relative;z-index:1;color:#fff;}}
.htitle .t1{{font-size:29px;font-weight:900;letter-spacing:1px;text-shadow:0 2px 4px rgba(26,58,107,.25);}}
.htitle .t2{{font-size:12px;font-weight:800;margin-top:1mm;background:rgba(255,255,255,.30);display:inline-block;padding:1mm 4mm;border-radius:20px;}}
.hmeta{{flex:0 0 76mm;position:relative;z-index:1;background:rgba(255,255,255,.94);border-radius:12px;padding:3mm 4mm;display:flex;flex-direction:column;gap:2mm;}}
.hmeta .r{{display:flex;align-items:center;gap:2mm;font-size:14px;font-weight:800;}}
.hmeta label{{color:var(--sky-d);white-space:nowrap;}}.hmeta .wl{{flex:1;border-bottom:2px solid var(--line);height:6mm;}}
.hmeta .mon{{display:inline-block;min-width:14mm;border-bottom:2px solid var(--line);}}
.body{{padding:4mm 6mm 0;}}
.plan{{display:flex;gap:4mm;margin-bottom:3mm;}}
.plan .pc{{flex:1;background:var(--soft);border-radius:10px;padding:2.5mm 3.5mm;border-left:5px solid var(--sky);}}
.plan .pc.b{{border-left-color:var(--orange);}}.plan .pc.c{{border-left-color:var(--pink);}}
.plan .ph{{font-size:11.5px;font-weight:900;margin-bottom:1mm;}}
.plan .pc .ph{{color:var(--sky-d);}}.plan .pc.b .ph{{color:var(--orange);}}.plan .pc.c .ph{{color:var(--pink);}}
.plan .psub{{font-size:8.5px;color:#9fb0c0;font-weight:700;}}
.plan .pl{{border-bottom:1.6px dotted #b6c6d4;height:6mm;margin-top:1mm;}}
.gtitle{{display:flex;align-items:center;gap:3mm;margin:1mm 0 2.5mm;}}
.gtitle .gt1{{font-size:16px;font-weight:900;color:var(--pink);}}
.gtitle .gt2{{font-size:10.5px;font-weight:800;color:var(--sky-d);background:rgba(78,197,224,.15);padding:1mm 3mm;border-radius:7px;}}
.gauges{{display:flex;flex-direction:column;gap:2.2mm;}}
.gauge{{display:flex;align-items:center;gap:2.5mm;padding:2mm 3mm;border:1.8px solid #e1e9f0;border-radius:11px;background:#fff;}}
.badge{{flex:0 0 9mm;height:9mm;border-radius:50%;display:flex;align-items:center;justify-content:center;
font-size:15px;font-weight:900;color:#fff;}}
.b-sky{{background:var(--sky);}}.b-pink{{background:var(--pink-l);}}.b-pur{{background:var(--purple);}}
.b-grn{{background:var(--green);}}.b-yel{{background:var(--yellow);color:var(--navy);}}.b-org{{background:var(--orange);}}
.gname{{flex:0 0 42mm;}}.gname .n1{{font-size:13px;font-weight:900;line-height:1.15;}}
.gname .n2{{font-size:8.5px;color:#94a3b8;font-weight:700;}}
.gtrack{{flex:1;display:flex;align-items:center;gap:1.2mm;}}
.slot{{flex:1;aspect-ratio:1;max-width:9mm;border:1.6px dashed #c7d3df;border-radius:50%;display:flex;
align-items:center;justify-content:center;font-size:8px;color:#cdd9e3;font-weight:800;}}
.slot.m{{border-color:var(--yellow);border-style:solid;background:rgba(255,201,60,.16);}}
.slot.goal{{border-color:var(--pink);border-style:solid;background:rgba(232,25,125,.10);}}
.st{{width:62%;height:62%;background:var(--yellow);
clip-path:polygon(50% 0,61% 35%,98% 35%,68% 57%,79% 91%,50% 70%,21% 91%,32% 57%,2% 35%,39% 35%);}}
.slot.goal .st{{background:var(--pink);}}
.c-sky{{border-color:var(--sky);}}.c-pink{{border-color:var(--pink-l);}}.c-pur{{border-color:var(--purple);}}
.c-grn{{border-color:var(--green);}}.c-yel{{border-color:var(--yellow);}}.c-org{{border-color:var(--orange);}}
.bottom{{display:flex;gap:4mm;margin-top:3.5mm;}}
.panel{{border:1.8px solid var(--sky);border-radius:11px;padding:3mm 4mm;background:#fff;}}
.panel h3{{font-size:12px;font-weight:900;margin-bottom:1.5mm;}}
.p-leg{{flex:0 0 56mm;border-color:#d6e3ec;background:var(--soft);}}.p-leg h3{{color:#5b7187;}}
.demo{{display:flex;align-items:center;gap:2mm;font-size:10px;font-weight:800;margin-bottom:1.5mm;}}
.demo .bar{{display:flex;gap:1mm;}}.demo .d{{width:5mm;height:5mm;border-radius:50%;border:1.5px dashed #c7d3df;}}
.demo .d.f{{border:none;background:var(--yellow);}}.demo.good .lab{{color:var(--green-d);}}.demo.few .lab{{color:var(--orange);}}
.p-rwd{{flex:0 0 80mm;border-color:var(--yellow);background:#fffdf5;}}.p-rwd h3{{color:var(--orange);}}
.p-rwd .rh{{display:flex;justify-content:space-between;align-items:baseline;}}
.p-rwd .tot{{font-size:9px;color:var(--sky-d);font-weight:800;}}
.p-rwd .tier{{display:flex;align-items:center;gap:2mm;margin-bottom:1.5mm;}}
.p-rwd .mb{{flex:0 0 auto;width:5mm;height:5mm;border-radius:50%;color:#fff;font-size:9px;font-weight:900;
display:flex;align-items:center;justify-content:center;}}
.p-rwd .cond{{flex:0 0 26mm;font-size:10px;font-weight:900;color:var(--orange);}}
.p-rwd .arrow{{color:#cbb36a;font-weight:900;}}.p-rwd .wbox{{flex:1;border-bottom:2px solid var(--line);height:5.5mm;}}
.p-rwd .tip{{font-size:8.5px;color:#b09a4e;font-weight:700;margin-top:.5mm;}}
.p-msg{{flex:1 1 auto;border-color:var(--pink-l);}}.p-msg h3{{color:var(--pink);}}
.p-msg .ln{{border-bottom:1.6px dotted #e3b9cf;height:7.5mm;}}
.p-msg .sign{{text-align:right;font-size:10px;font-weight:800;color:#9fb0c0;margin-top:1mm;}}
.footnote{{margin:3mm 6mm 0;font-size:9px;color:#7c8da0;font-weight:700;line-height:1.55;border-top:1.6px dashed #d6e3ec;
padding-top:2mm;display:flex;gap:4mm;align-items:flex-start;}}
.footnote .use{{flex:1;}}.footnote b{{color:var(--sky-d);}}
.footnote .brand{{flex:0 0 auto;text-align:right;color:var(--sky-d);font-weight:900;white-space:nowrap;}}
.footnote .brand small{{display:block;color:#9fb0c0;font-weight:700;letter-spacing:2px;}}
@media print{{body{{background:#fff;}}.screenbar{{display:none;}}.sheet{{box-shadow:none;margin:0;}}}}
</style></head><body>
<div class="screenbar">
  <h1>わたしの がんばりカード（夢門塾・1ヶ月）</h1>
  <p>「印刷」→ <b>A4・横向き</b>。1シート＝1ヶ月。毎日できた目標にシールを貼り、月末にゲージで得意・苦手を確認します。</p>
  <button class="pbtn" onclick="window.print()">印刷する</button>
</div>
<div class="sheet">
  <div class="brandhead">
    <span class="cloud c1"></span><span class="cloud c2"></span><span class="cloud c3"></span>
    <div class="logo"><img src="{LOGO}" alt="夢門塾"></div>
    <div class="htitle">
      <div class="t1">わたしの がんばりカード</div>
      <div class="t2">1ヶ月の せいちょうきろく ・ まいにち ふりかえり</div>
    </div>
    <div class="hmeta">
      <div class="r"><label>なまえ</label><span class="wl"></span></div>
      <div class="r"><label>き　かん</label><span class="mon"></span><span>月</span></div>
    </div>
  </div>
  <div class="body">
    <div class="plan">
      <div class="pc"><div class="ph">ながい もくひょう <span class="psub">（個別支援計画より）</span></div><div class="pl"></div></div>
      <div class="pc b"><div class="ph">今月の チャレンジ <span class="psub">（自分で きめる 近い 目標）</span></div><div class="pl"></div></div>
      <div class="pc c"><div class="ph">おうえん <span class="psub">（おうちの人・先生から）</span></div><div class="pl"></div></div>
    </div>
    <div class="gtitle">
      <span class="gt1">できた日に シールを はろう（1ヶ月）</span>
      <span class="gt2">シールが ふえると せいちょうが 見える！</span>
    </div>
    <div class="gauges">{gauges_html}
    </div>
    <div class="bottom">
      <div class="panel p-leg">
        <h3>シールの 見かた</h3>
        <div class="demo good"><span class="bar"><span class="d f"></span><span class="d f"></span><span class="d f"></span><span class="d f"></span><span class="d f"></span></span><span class="lab">いっぱい ＝ とくい！</span></div>
        <div class="demo few"><span class="bar"><span class="d f"></span><span class="d"></span><span class="d"></span><span class="d"></span><span class="d"></span></span><span class="lab">すくない ＝ いっしょに れんしゅう</span></div>
      </div>
      <div class="panel p-rwd">
        <div class="rh"><h3>今月の ごほうび</h3><span class="tot">シール合計（　）こ</span></div>
        <div class="tier"><span class="mb" style="background:#CD7F32">1</span><span class="cond">シール 20こ</span><span class="arrow">→</span><span class="wbox"></span></div>
        <div class="tier"><span class="mb" style="background:#A8B0B8">2</span><span class="cond">ゲージ 3本 クリア</span><span class="arrow">→</span><span class="wbox"></span></div>
        <div class="tier"><span class="mb" style="background:#D4A017">3</span><span class="cond">ぜんぶ クリア</span><span class="arrow">→</span><span class="wbox"></span></div>
        <div class="tip">＊ごほうびは 子どもと いっしょに きめると やる気アップ！</div>
      </div>
      <div class="panel p-msg">
        <h3>今月の ふりかえり・がんばったね メッセージ</h3>
        <div class="ln"></div><div class="ln"></div>
        <div class="sign">先生・おうちの人 ＿＿＿＿＿＿＿</div>
      </div>
    </div>
  </div>
  <div class="footnote">
    <div class="use"><b>使い方</b>　1ヶ月を 1まいで 記録。毎日の 終わりに 子どもと いっしょに ふり返り、できた 目標に シールを 1まい 貼る（注意の 代わりに「できた所を 見つけて すぐ ほめる」）。
      月末に ゲージを 見て、得意・苦手を つぎの 月の 目標に いかす。目標①〜④は 個別支援計画の 課題から、⑤⑥は 成長・チャレンジから 書きかえてOK。</div>
    <div class="brand">放課後等デイサービス 夢門塾<small>MUMONJUKU</small></div>
  </div>
</div></body></html>'''
open(os.path.join(HERE,'ganbari-card-mumonjuku.html'),'w').write(HTML)
print("html written", len(HTML))
