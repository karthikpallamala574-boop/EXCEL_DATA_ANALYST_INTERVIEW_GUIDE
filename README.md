<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Excel for Data Analyst — Interview Guide</title>
<link rel="preconnect" href="https://fonts.googleapis.com">
<link href="https://fonts.googleapis.com/css2?family=Fraunces:opsz,wght@9..144,400;9..144,600;9..144,700&family=IBM+Plex+Mono:wght@400;500;600&display=swap" rel="stylesheet">
<style>
  :root{
    --paper:#F5F3EC; --paper-2:#EFEDE3; --ink:#17231C; --ink-soft:#4A5A50;
    --green-deep:#173B2C; --green-mid:#2D6A4F; --line:#C9D1C6; --gold:#B8912B;
    box-sizing:border-box;
    padding-top:env(safe-area-inset-top,0px);
    padding-bottom:env(safe-area-inset-bottom,0px);
  }
  @media (prefers-color-scheme: dark){
    :root:not([data-theme="light"]){
      --paper:#12160F; --paper-2:#171C13; --ink:#EAE7DC; --ink-soft:#A7B3A9;
      --green-deep:#8FD9B6; --green-mid:#6FBE96; --line:#31392E; --gold:#D9B65B;
    }
  }
  :root[data-theme="dark"]{
    --paper:#12160F; --paper-2:#171C13; --ink:#EAE7DC; --ink-soft:#A7B3A9;
    --green-deep:#8FD9B6; --green-mid:#6FBE96; --line:#31392E; --gold:#D9B65B;
  }
  *{box-sizing:border-box;}
  html{scroll-padding-top:env(safe-area-inset-top,0px);}
  html,body{height:100%;}
  body{
    margin:0; background:var(--paper); color:var(--ink);
    font-family:-apple-system,BlinkMacSystemFont,"Segoe UI",sans-serif;
    -webkit-font-smoothing:antialiased; overflow-x:hidden;
  }
  .wrap{max-width:720px; margin:0 auto; padding:0 24px;}

  /* Hero */
  header{
    min-height:78vh; display:flex; flex-direction:column; justify-content:center;
    border-bottom:1px solid var(--line); position:relative;
  }
  .kicker{font-family:"IBM Plex Mono",monospace; font-size:13px; color:var(--green-mid); letter-spacing:.02em;}
  h1{
    font-family:"Fraunces",serif; font-weight:600; font-size:clamp(38px,7vw,64px);
    line-height:1.02; margin:14px 0 18px; max-width:11ch;
  }
  header p{font-size:18px; line-height:1.55; color:var(--ink-soft); max-width:52ch;}
  .grid-ref{
    position:absolute; right:24px; top:24px; font-family:"IBM Plex Mono",monospace;
    font-size:12px; color:var(--ink-soft); text-align:right; line-height:1.6;
  }

  /* Format strip */
  .format{
    display:flex; flex-wrap:wrap; gap:0; border-bottom:1px solid var(--line);
    font-family:"IBM Plex Mono",monospace; font-size:12.5px; color:var(--ink-soft);
  }
  .format span{
    padding:14px 16px; border-right:1px solid var(--line);
  }
  .format span:last-child{border-right:none;}

  /* Scroll stage */
  #stage{ perspective:1200px; padding:8vh 0 12vh; }
  .row{
    display:grid; grid-template-columns:88px 1fr; gap:22px;
    padding:34px 0; border-bottom:1px solid var(--line);
    transform-style:preserve-3d; will-change:transform, opacity;
    transition:transform .05s linear;
  }
  .row .num{
    font-family:"IBM Plex Mono",monospace; font-size:15px; color:var(--gold);
    padding-top:4px;
  }
  .row .num small{display:block; color:var(--ink-soft); font-size:11px; margin-top:4px;}
  .row h2{
    font-family:"Fraunces",serif; font-weight:600; font-size:24px; margin:0 0 8px;
  }
  .row p{margin:0; color:var(--ink-soft); font-size:15px; line-height:1.55;}

  /* About / footer */
  footer{border-top:1px solid var(--line); padding:56px 0 72px;}
  footer h3{font-family:"Fraunces",serif; font-size:26px; margin:0 0 10px;}
  footer p{color:var(--ink-soft); font-size:15px; max-width:52ch; line-height:1.6;}
  .links{display:flex; flex-wrap:wrap; gap:10px; margin-top:20px; font-family:"IBM Plex Mono",monospace; font-size:13px;}
  .links a{
    color:var(--green-deep); text-decoration:none; border:1px solid var(--line);
    padding:9px 14px; border-radius:2px;
  }
  .links a:hover{border-color:var(--green-mid);}
  .star{margin-top:28px; font-size:13px; color:var(--ink-soft);}

  @media (max-width:560px){
    .row{grid-template-columns:56px 1fr; gap:14px;}
    .grid-ref{display:none;}
  }
</style>
</head>
<body>

<header class="wrap">
  <div class="kicker">README · A1:B12</div>
  <h1>Excel for Data Analyst</h1>
  <p>An 11-level interview training guide — from SUM and COUNT up through Power Query and VBA. Every concept: definition, example, spoken interview answer, and the follow-ups that trip people up.</p>
  <div class="grid-ref">fx&nbsp;&nbsp;=SCROLL(levels)</div>
</header>

<div class="format wrap">
  <span>Definition</span><span>Explanation</span><span>Example</span><span>Interview Answer</span><span>Follow-ups</span><span>Common mistakes</span>
</div>

<div id="stage" class="wrap">
  <div class="row" data-i="1"><div class="num">01<small>A1</small></div><div><h2>Excel Basics</h2><p>SUM, AVERAGE, MIN, MAX, COUNT, COUNTA, COUNTBLANK — the functions every other level assumes you already know cold.</p></div></div>
  <div class="row" data-i="2"><div class="num">02<small>B1</small></div><div><h2>Conditional Functions</h2><p>COUNTIF, SUMIF, AVERAGEIF and their multi-criteria IFS siblings, plus MAXIFS and MINIFS.</p></div></div>
  <div class="row" data-i="3"><div class="num">03<small>C1</small></div><div><h2>Text &amp; Data Cleaning</h2><p>UPPER, LOWER, PROPER, LEFT, RIGHT, MID, LEN, CONCAT, FIND, SEARCH, REPLACE, SUBSTITUTE, TRIM, CLEAN.</p></div></div>
  <div class="row" data-i="4"><div class="num">04<small>D1</small></div><div><h2>References</h2><p>Relative, absolute, and mixed references — the difference that breaks copied formulas.</p></div></div>
  <div class="row" data-i="5"><div class="num">05<small>E1</small></div><div><h2>Math &amp; Date</h2><p>ROUND family, ABS, POWER, SQRT, TODAY, NOW, date-part functions, WEEKDAY, NETWORKDAYS, WORKDAY.</p></div></div>
  <div class="row" data-i="6"><div class="num">06<small>F1</small></div><div><h2>Data Handling</h2><p>Paste Special, Flash Fill, Fill Series, Text to Columns, Sort, Filter, Data Validation, Conditional Formatting.</p></div></div>
  <div class="row" data-i="7"><div class="num">07<small>G1</small></div><div><h2>Logical Functions</h2><p>IF, AND, OR, NOT, and nested IF — reasoning through branching conditions out loud.</p></div></div>
  <div class="row" data-i="8"><div class="num">08<small>H1</small></div><div><h2>Lookup</h2><p>VLOOKUP, XLOOKUP, INDEX, MATCH, and INDEX+MATCH — the question every interview asks.</p></div></div>
  <div class="row" data-i="9"><div class="num">09<small>I1</small></div><div><h2>Analysis</h2><p>SUBTOTAL, Pivot Tables, Pivot Charts — turning raw rows into a summary someone can act on.</p></div></div>
  <div class="row" data-i="10"><div class="num">10<small>J1</small></div><div><h2>Data Analyst Tools</h2><p>Power Query, data cleaning pipelines, and data modeling for multi-table analysis.</p></div></div>
  <div class="row" data-i="11"><div class="num">11<small>K1</small></div><div><h2>Advanced Excel</h2><p>VBA and consolidation — where Excel stops being a spreadsheet and starts being a tool.</p></div></div>
</div>

<footer class="wrap">
  <h3>About</h3>
  <p>Built by Karthik Pallamala (Jeeva), an entry-level Data Analyst working in Power BI, DAX, SQL, and Excel — for freshers preparing for their first Data Analyst or Business Analyst interview.</p>
  <div class="links">
    <a href="https://karthikpallamala574-boop.github.io/Portfolio" target="_blank" rel="noopener">Portfolio</a>
    <a href="https://www.linkedin.com/in/karthik-pallamala-aa1175382/" target="_blank" rel="noopener">LinkedIn</a>
    <a href="https://github.com/karthikpallamala574-boop" target="_blank" rel="noopener">GitHub</a>
  </div>
  <p class="star">⭐ Free to use and share for personal learning and interview prep.</p>
</footer>

<script>
(function(){
  var rows = Array.prototype.slice.call(document.querySelectorAll('.row'));
  var reduce = window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  function update(){
    if(reduce) return;
    var vh = window.innerHeight;
    rows.forEach(function(r){
      var rect = r.getBoundingClientRect();
      var center = rect.top + rect.height/2;
      var dist = (center - vh/2) / vh; // -1..1 roughly
      var clamped = Math.max(-1, Math.min(1, dist));
      var rotate = clamped * 10;
      var z = -Math.abs(clamped) * 70;
      var opacity = 1 - Math.min(0.55, Math.abs(clamped) * 0.5);
      r.style.transform = 'rotateX(' + rotate + 'deg) translateZ(' + z + 'px)';
      r.style.opacity = opacity;
    });
  }
  document.addEventListener('scroll', update, {passive:true});
  window.addEventListener('resize', update);
  update();
})();
</script>
</body>
</html>
