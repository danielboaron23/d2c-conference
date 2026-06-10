#!/usr/bin/env python3
"""
Build workshop-opening.html — a trimmed copy of the conference deck
(index-v4.html) used as the OPENING of the Claude Code terminal workshop.

Transforms applied (all reproducible — re-run any time):
  - Remove the data slides: 3a, 3b, 3c, 3d, 13.
    (13 = the conference Sponsors & Thanks slide; not part of the workshop opening.)
  - Null-guard the reset block that references those removed slides
    (s3b / s3d) so navigation never throws.
  - Inject the workshop shell bridge (postMessage nav + boundary cross).

Lives in d2c-conference so the deck's local assets/ + fonts/ resolve.
Run: python3 build-workshop-opening.py
"""
import re, os

HERE = os.path.dirname(os.path.abspath(__file__))
SRC = os.path.join(HERE, "index-v4.html")
OUT = os.path.join(HERE, "workshop-opening.html")
REMOVE = {"3a", "3b", "3c", "3d", "13"}

BRIDGE = r'''<!-- ===== workshop shell bridge (injected) ===== -->
<script>
(function(){
  function ds(){return document.querySelector('deck-stage');}
  function info(){
    var d=ds();
    if(d){
      var secs=Array.prototype.slice.call(d.children).filter(function(s){return s.tagName==='SECTION';});
      var cur=secs.findIndex(function(s){return s.hasAttribute('data-deck-active');});
      return {engine:'deck',cur:cur,total:secs.length};
    }
    var sl=Array.prototype.slice.call(document.querySelectorAll('.slide'));
    var cur=sl.findIndex(function(s){return s.classList.contains('active');});
    return {engine:'v4',cur:cur,total:sl.length};
  }
  var NEXT={ArrowRight:1,ArrowDown:1,PageDown:1,' ':1,Spacebar:1,Enter:1};
  var PREV={ArrowLeft:1,ArrowUp:1,PageUp:1};
  function labels(){
    var d=ds();
    if(d){return Array.prototype.slice.call(d.children).filter(function(s){return s.tagName==='SECTION';}).map(function(s){return s.getAttribute('data-label')||'';});}
    return Array.prototype.slice.call(document.querySelectorAll('.slide')).map(function(s){return s.getAttribute('data-slide')||'';});
  }
  function report(){var i=info();try{parent.postMessage({type:'ws-pos',engine:i.engine,cur:i.cur,total:i.total,labels:labels()},'*');}catch(e){}}
  function advance(dir){
    var d=ds();
    if(d){ if(dir==='next')d.next(); else d.prev(); }
    else { var b=document.getElementById(dir==='next'?'next-btn':'prev-btn'); if(b)b.click(); }
    setTimeout(report,60);
  }
  function goTo(index){
    var d=ds();
    if(d){ if(typeof d.goTo==='function') d.goTo(index); }
    else { var dots=document.querySelectorAll('.nav-dot'); if(dots && dots[index]) dots[index].click(); }
    setTimeout(report,60);
  }
  window.addEventListener('keydown',function(e){
    if((e.metaKey||e.ctrlKey)&&(e.key==='k'||e.key==='K')){e.preventDefault();e.stopImmediatePropagation();try{parent.postMessage({type:'ws-hotkey',key:'toc'},'*');}catch(_){}return;}
    if(e.key==='m'||e.key==='M'){e.preventDefault();e.stopImmediatePropagation();try{parent.postMessage({type:'ws-hotkey',key:'toc'},'*');}catch(_){}return;}
    var i=info(); if(i.cur<0){return;}
    if(NEXT[e.key] && i.cur===i.total-1){e.preventDefault();e.stopImmediatePropagation();try{parent.postMessage({type:'ws-cross',dir:'next'},'*');}catch(_){}return;}
    if(PREV[e.key] && i.cur===0){e.preventDefault();e.stopImmediatePropagation();try{parent.postMessage({type:'ws-cross',dir:'prev'},'*');}catch(_){}return;}
    setTimeout(report,40);
  },true);
  window.addEventListener('message',function(ev){
    var m=ev.data||{}; if(!m||m.type!=='ws-cmd'){return;}
    if(m.cmd==='report'){report();}
    else if(m.cmd==='focus'){try{window.focus();}catch(e){}}
    else if(m.cmd==='next'||m.cmd==='prev'){advance(m.cmd);}
    else if(m.cmd==='goto'){goTo(m.index|0);}
  });
  window.addEventListener('load',function(){try{window.focus();}catch(e){} setTimeout(report,150);});
  setTimeout(report,300);
})();
</script>'''

def match_div_end(text, start):
    i, depth = start, 0
    tag_re = re.compile(r'<(/?)div\b[^>]*?(/?)>', re.I)
    while True:
        m = tag_re.search(text, i)
        if not m:
            return -1
        if m.group(1) == '/':
            depth -= 1
        elif m.group(2) != '/':
            depth += 1
        i = m.end()
        if depth == 0:
            return i

def main():
    src = open(SRC, encoding="utf-8").read()

    # 1) remove data slides (back-to-front to keep indices valid)
    opens = [(m.group(1), m.start())
             for m in re.finditer(r'<div class="slide" data-slide="([^"]+)">', src)]
    removals = []
    for sid, st in opens:
        if sid in REMOVE:
            en = match_div_end(src, st)
            assert en != -1, f"no end for slide {sid}"
            removals.append((st, en))
    out = src
    for st, en in sorted(removals, key=lambda r: -r[0]):
        line_start = out.rfind("\n", 0, st) + 1
        if out[en:en + 1] == "\n":
            en += 1
        out = out[:line_start] + out[en:]

    remaining = re.findall(r'<div class="slide" data-slide="([^"]+)">', out)
    assert not (set(remaining) & REMOVE), "data slides still present"

    # 2) null-guard the reset block (references removed slides 3b / 3d)
    out = out.replace(
        '''    var s3b = document.querySelector('[data-slide="3b"]');
    s3b.querySelectorAll('.sd-num-value').forEach(function(el, i) {''',
        '''    var s3b = document.querySelector('[data-slide="3b"]');
    if (s3b) s3b.querySelectorAll('.sd-num-value').forEach(function(el, i) {''', 1)
    out = out.replace(
        '''    var s3d = document.querySelector('[data-slide="3d"]');
    var s3dStats = s3d.querySelectorAll('.sd-hire-stat');
    var s3dOriginals = ['73%', '68%', '$294K', '96%'];
    s3dStats.forEach(function(el, i) { el.textContent = s3dOriginals[i]; });''',
        '''    var s3d = document.querySelector('[data-slide="3d"]');
    if (s3d) {
      var s3dStats = s3d.querySelectorAll('.sd-hire-stat');
      var s3dOriginals = ['73%', '68%', '$294K', '96%'];
      s3dStats.forEach(function(el, i) { el.textContent = s3dOriginals[i]; });
    }''', 1)

    # 3) inject shell bridge
    out = out.replace("</body>", BRIDGE + "\n</body>", 1)

    open(OUT, "w", encoding="utf-8").write(out)
    print(f"Built {OUT}")
    print(f"  removed slides: {sorted(REMOVE)}")
    print(f"  remaining slides: {len(remaining)} -> {remaining}")
    assert "if (s3b)" in out and "if (s3d)" in out, "guards not applied"
    assert "ws-cross" in out, "bridge not injected"
    print("  guards + bridge: OK")

if __name__ == "__main__":
    main()
