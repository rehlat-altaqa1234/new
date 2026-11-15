/* PowerPoint-faithful slide engine: renders window.DECK on a 16:9 stage and replays
   its animations, triggers, hyperlinks and media. */
(function () {
  'use strict';
  const D = window.DECK, W = D.w, H = D.h;
  const stage = document.getElementById('stage');
  const wrap = document.getElementById('wrap');
  const counter = document.getElementById('counter');
  let cur = -1, els = {}, shapesById = {}, timers = [], anims = [], step = 0, trigState = {}, audios = {};
  const PT = 96 / 72;
  const FONT = "'Sakkal Majalla','Majalla','SMFallback','Almarai',sans-serif";

  /* ---------------- scaling */
  function fit() {
    const vw = window.innerWidth, vh = window.innerHeight;
    const s = Math.min(vw / W, vh / H);
    stage.style.transform = `translate(${(vw - W * s) / 2}px,${(vh - H * s) / 2}px) scale(${s})`;
  }
  window.addEventListener('resize', fit);

  /* ---------------- geometry */
  function svgPath(g, w, h, av) {
    const a = (n, d) => (av && av[n]) ? parseInt(av[n].split(' ')[1]) : d;
    switch (g) {
      case 'triangle': return `M${w / 2},0 L${w},${h} L0,${h} Z`;
      case 'hexagon': { const o = Math.min(w, h) * 0.25; return `M${o},0 L${w - o},0 L${w},${h / 2} L${w - o},${h} L${o},${h} L0,${h / 2} Z`; }
      case 'leftArrow': { const hw = Math.min(w, h) * 0.5, t = h * 0.25; return `M0,${h / 2} L${hw},0 L${hw},${t} L${w},${t} L${w},${h - t} L${hw},${h - t} L${hw},${h} Z`; }
      case 'rightArrow': { const hw = Math.min(w, h) * 0.5, t = h * 0.25; return `M${w},${h / 2} L${w - hw},0 L${w - hw},${t} L0,${t} L0,${h - t} L${w - hw},${h - t} L${w - hw},${h} Z`; }
      case 'downArrow': { const hh = Math.min(w, h) * 0.5, t = w * 0.25; return `M${w / 2},${h} L0,${h - hh} L${t},${h - hh} L${t},0 L${w - t},0 L${w - t},${h - hh} L${w},${h - hh} Z`; }
      case 'upArrow': { const hh = Math.min(w, h) * 0.5, t = w * 0.25; return `M${w / 2},0 L0,${hh} L${t},${hh} L${t},${h} L${w - t},${h} L${w - t},${hh} L${w},${hh} Z`; }
      case 'pie': {
        const s = a('adj1', 0) / 60000 * Math.PI / 180, e = a('adj2', 16200000) / 60000 * Math.PI / 180;
        const cx = w / 2, cy = h / 2, rx = w / 2, ry = h / 2;
        let sweep = e - s; if (sweep <= 0) sweep += 2 * Math.PI;
        const x1 = cx + rx * Math.cos(s), y1 = cy + ry * Math.sin(s), x2 = cx + rx * Math.cos(s + sweep), y2 = cy + ry * Math.sin(s + sweep);
        return `M${cx},${cy} L${x1},${y1} A${rx},${ry} 0 ${sweep > Math.PI ? 1 : 0} 1 ${x2},${y2} Z`;
      }
      case 'sun': {
        let d = ''; const cx = w / 2, cy = h / 2, r = Math.min(w, h) * 0.25;
        d += `M${cx + r},${cy} A${r},${r} 0 1 1 ${cx - r},${cy} A${r},${r} 0 1 1 ${cx + r},${cy} Z`;
        for (let i = 0; i < 8; i++) {
          const an = i * Math.PI / 4, r1 = Math.min(w, h) * 0.34, r2 = Math.min(w, h) * 0.5, sp = 0.18;
          d += `M${cx + r1 * Math.cos(an - sp)},${cy + r1 * Math.sin(an - sp)} L${cx + r2 * Math.cos(an)},${cy + r2 * Math.sin(an)} L${cx + r1 * Math.cos(an + sp)},${cy + r1 * Math.sin(an + sp)} Z`;
        }
        return d;
      }
      case 'wave': { const k = h * 0.13; return `M0,${k} C${w / 3},${-k} ${w * 2 / 3},${3 * k} ${w},${k} L${w},${h - k} C${w * 2 / 3},${h + k} ${w / 3},${h - 3 * k} 0,${h - k} Z`; }
      case 'can': { const e = Math.min(h * 0.25, w * 0.25); return `M0,${e / 2} A${w / 2},${e / 2} 0 0 1 ${w},${e / 2} L${w},${h - e / 2} A${w / 2},${e / 2} 0 0 1 0,${h - e / 2} Z M0,${e / 2} A${w / 2},${e / 2} 0 0 0 ${w},${e / 2}`; }
      default: return null;
    }
  }

  /* ---------------- rendering */
  function el(tag, cls, parent) { const e = document.createElement(tag); if (cls) e.className = cls; if (parent) parent.appendChild(e); return e; }

  function renderText(tx, box) {
    const t = el('div', 'tx', box);
    const [ti, ri, bi, li] = tx.ins;
    t.style.padding = `${ti}px ${ri}px ${bi}px ${li}px`;
    t.style.justifyContent = { t: 'flex-start', ctr: 'center', b: 'flex-end' }[tx.anchor] || 'flex-start';
    if (!tx.wrap) t.style.whiteSpace = 'nowrap';
    const sc = tx.scale || 1;
    for (const p of tx.paras) {
      const pe = el('div', 'p', t);
      if (p.rtl) pe.dir = 'rtl';
      pe.style.textAlign = { l: 'left', r: 'right', ctr: 'center', just: 'justify', dist: 'justify' }[p.al || 'l'];
      const ls = (p.ls || 1) * 1.2;
      pe.style.lineHeight = ls;
      if (p.sb) pe.style.marginTop = p.sb * PT + 'px';
      if (p.sa) pe.style.marginBottom = p.sa * PT + 'px';
      let maxsz = 0;
      if (p.bu) { const b = el('span', null, pe); b.textContent = p.bu + ' '; }
      for (const r of p.runs) {
        if (r.br) { el('br', null, pe); continue; }
        const s = el('span', null, pe);
        s.textContent = r.t;
        const sz = (r.sz || 18) * sc; maxsz = Math.max(maxsz, sz);
        s.style.fontSize = sz * PT + 'px';
        if (r.b) s.style.fontWeight = '700';
        if (r.i) s.style.fontStyle = 'italic';
        if (r.u) s.style.textDecoration = 'underline';
        s.style.color = r.c || '#000';
        if (r.bl) { s.style.verticalAlign = r.bl > 0 ? 'super' : 'sub'; s.style.fontSize = sz * PT * 0.7 + 'px'; }
        if (r.link) bindLink(s, r.link);
      }
      if (!p.runs.length || p.runs.every(r => r.br || !r.t)) {
        pe.style.fontSize = ((p.esz || maxsz || 18) * sc * PT) + 'px';
        if (!p.runs.length) pe.innerHTML = '&#8203;';
      }
    }
  }

  function bindLink(node, lk) {
    node.classList.add('clickable');
    node.addEventListener('click', ev => {
      ev.stopPropagation();
      if (lk.slide != null) go(lk.slide);
      else if (lk.url) { const a = document.createElement('a'); a.href = lk.url; a.target = '_blank'; a.rel = 'noopener'; document.body.appendChild(a); a.click(); a.remove(); }
      else if (lk.jump === 'nextslide') go(cur + 1);
      else if (lk.jump === 'previousslide') go(cur - 1);
      else if (lk.jump === 'firstslide') go(0);
      else if (lk.jump === 'lastslide') go(D.slides.length - 1);
    });
  }

  function render(s, parent) {
    const o = el('div', 'shp', parent);           // outer: position + motion paths + visibility
    o.style.left = s.x + 'px'; o.style.top = s.y + 'px';
    o.style.width = s.w + 'px'; o.style.height = s.h + 'px';
    const inn = el('div', 'inn', o);              // inner: rotation, entrance/emphasis effects
    const base = `rotate(${s.rot || 0}deg)` + (s.fh ? ' scaleX(-1)' : '') + (s.fv ? ' scaleY(-1)' : '');
    inn.style.transform = base; inn.dataset.base = base;
    els[s.id] = o; shapesById[s.id] = s; o._inn = inn; o._s = s;
    if (s.k === 'grp') {
      for (const k of s.kids) render(k, inn);
    } else if (s.k === 'pic') {
      if (s.mk === 'video') {
        const v = el('video', 'vid', inn); v.src = s.media; v.controls = true; v.playsInline = true; v.preload = 'metadata';
        if (s.src) v.poster = s.src;
        v.addEventListener('click', e => e.stopPropagation());
        audios[s.id] = v; o.dataset.media = '1';
      } else {
        if (s.media) { const a = new Audio(s.media); a.preload = 'auto'; audios[s.id] = a; }
        if (s.src) {
          const box = el('div', 'imgbox', inn);
          const im = el('img', null, box); im.src = s.src; im.alt = s.name || ''; im.draggable = false;
          if (s.crop) {
            const [t, r, b, l] = s.crop.map(v => v / 100);
            const iw = s.w / (1 - l - r), ih = s.h / (1 - t - b);
            Object.assign(im.style, { width: iw + 'px', height: ih + 'px', left: -l * iw + 'px', top: -t * ih + 'px' });
          }
        }
      }
      if (s.sh) inn.style.filter = `drop-shadow(${s.sh})`;
    } else {
      const path = svgPath(s.g, s.w, s.h, s.av);
      if (path) {
        const ns = 'http://www.w3.org/2000/svg';
        const svg = document.createElementNS(ns, 'svg');
        svg.setAttribute('viewBox', `0 0 ${s.w} ${s.h}`); svg.setAttribute('class', 'geo');
        const pe = document.createElementNS(ns, 'path'); pe.setAttribute('d', path);
        pe.setAttribute('fill', s.fill && !s.fill.startsWith('linear') ? s.fill : 'none');
        if (s.ln) { pe.setAttribute('stroke', s.ln.c); pe.setAttribute('stroke-width', s.ln.w); }
        svg.appendChild(pe); inn.appendChild(svg);
        if (s.sh) svg.style.filter = `drop-shadow(${s.sh})`;
      } else {
        const bg = el('div', 'geo', inn);
        if (s.fill) bg.style.background = s.fill;
        if (s.g === 'ellipse') bg.style.borderRadius = '50%';
        if (s.g === 'roundRect') {
          const adj = s.av && s.av.adj ? parseInt(s.av.adj.split(' ')[1]) / 100000 : 0.16667;
          bg.style.borderRadius = Math.min(s.w, s.h) * adj + 'px';
        }
        if (s.sh && (s.fill || s.ln)) bg.style.boxShadow = s.sh;
        if (s.ln) {
          const st = el('div', 'stroke', bg);
          const w = s.ln.w;
          Object.assign(st.style, { inset: -w / 2 + 'px', border: `${w}px ${s.ln.d ? 'dashed' : 'solid'} ${s.ln.c}`, borderRadius: bg.style.borderRadius ? `calc(${bg.style.borderRadius} + ${w / 2}px)` : '0' });
          if (s.g === 'ellipse') st.style.borderRadius = '50%';
        }
      }
      if (s.tx) renderText(s.tx, inn);
    }
    if (s.link) bindLink(o, s.link);
    return o;
  }

  /* ---------------- animation helpers */
  function later(fn, ms) { const t = setTimeout(fn, ms); timers.push(t); return t; }
  function track(a) { anims.push(a); return a; }
  function show(id, v) { const e = els[id]; if (e) e.style.visibility = v ? '' : 'hidden'; }

  function vars(id) {
    const s = shapesById[id];
    const g = absBox(id);
    return { ppt_x: (g.x + g.w / 2) / W, ppt_y: (g.y + g.h / 2) / H, ppt_w: g.w / W, ppt_h: g.h / H, s };
  }
  function absBox(id) {
    const e = els[id]; let x = 0, y = 0, n = e;
    while (n && n !== stageInner) { x += n.offsetLeft; y += n.offsetTop; n = n.offsetParent === stageInner ? null : n.offsetParent; }
    return { x, y, w: e.offsetWidth, h: e.offsetHeight };
  }
  function evalF(expr, v) {
    if (expr == null) return null;
    let s = String(expr).replace(/#ppt_x/g, v.ppt_x).replace(/#ppt_y/g, v.ppt_y).replace(/#ppt_w/g, v.ppt_w).replace(/#ppt_h/g, v.ppt_h);
    if (!/^[-+*/().\d\seE]*$/.test(s)) return null;
    try { return Function('return (' + s + ')')(); } catch (e) { return null; }
  }
  function sampleTav(tav, t, v) {
    if (!tav || !tav.length) return null;
    const pts = tav.map(([tm, val]) => [tm, evalF(val, v)]).filter(p => p[1] != null);
    if (!pts.length) return null;
    if (t <= pts[0][0]) return pts[0][1];
    for (let i = 1; i < pts.length; i++) {
      if (t <= pts[i][0]) { const [t0, a] = pts[i - 1], [t1, b] = pts[i]; return a + (b - a) * ((t - t0) / ((t1 - t0) || 1)); }
    }
    return pts[pts.length - 1][1];
  }
  function wipeClip(f, tr, p) {   // p: progress 0..1 of reveal
    const m = /wipe\((\w+)\)/.exec(f || ''); if (!m) return null;
    const q = (1 - p) * 100;
    return { down: `inset(0 0 ${q}% 0)`, up: `inset(${q}% 0 0 0)`, right: `inset(0 ${q}% 0 0)`, left: `inset(0 0 0 ${q}%)` }[m[1]] || null;
  }
  const num = (x, d) => x == null ? d : +x;

  function runEffect(e, instant) {
    const byId = {};
    for (const b of e.b) { if (b.id) (byId[b.id] = byId[b.id] || []).push(b); }
    let end = 0;
    for (const id in byId) {
      const target = els[id]; if (!target) continue;
      const bs = byId[id];
      const inn = target._inn, base = inn.dataset.base;
      // visibility sets
      for (const b of bs.filter(b => b.t === 'set' && b.attr === 'style.visibility')) {
        const d = num(b.delay, 0);
        if (instant) show(id, b.to === 'visible'); else later(() => show(id, b.to === 'visible'), d);
      }
      const isExit = e.cls === 'exit';
      if (instant) {
        if (e.cls === 'entr') show(id, true);
        if (isExit) show(id, false);
        continue;
      }
      // entrance / exit composite (fade / wipe / ppt_x,y,w,h formulas)
      const fx = bs.filter(b => b.t === 'animEffect' || (b.t === 'anim' && /^ppt_/.test(b.attr || '')));
      if (fx.length && (e.cls === 'entr' || isExit)) {
        const dur = Math.max(...fx.map(b => num(b.dur, 500))), d0 = Math.min(...fx.map(b => num(b.delay, 0)));
        const v = vars(id), N = 16, frames = [];
        for (let i = 0; i <= N; i++) {
          const t = i / N; let op = 1, dx = 0, dy = 0, sx = 1, sy = 1, clip = null;
          for (const b of fx) {
            if (b.t === 'animEffect') {
              const prog = b.tr === 'out' ? 1 - t : t;
              if (/fade|dissolve/.test(b.f || '')) op = prog;
              else { const c = wipeClip(b.f, b.tr, prog); if (c) clip = c; else op = prog; }
            } else {
              const val = sampleTav(b.tav, t, v);
              if (val == null) continue;
              if (b.attr === 'ppt_x') dx = (val - v.ppt_x) * W;
              if (b.attr === 'ppt_y') dy = (val - v.ppt_y) * H;
              if (b.attr === 'ppt_w') sx = v.ppt_w ? val / v.ppt_w : 1;
              if (b.attr === 'ppt_h') sy = v.ppt_h ? val / v.ppt_h : 1;
            }
          }
          const f = { offset: t, opacity: op, transform: `translate(${dx}px,${dy}px) scale(${sx},${sy}) ${base}` };
          f.clipPath = clip || 'inset(0 0 0 0)';
          frames.push(f);
        }
        later(() => {
          const a = track(inn.animate(frames, { duration: dur, easing: 'linear', fill: 'none' }));
          if (isExit) a.onfinish = () => show(id, false);
        }, d0);
        end = Math.max(end, d0 + dur);
      } else if (isExit) {
        later(() => show(id, false), 0);
      }
      // emphasis
      for (const b of bs) {
        const d = num(b.delay, 0), dur = num(b.dur, 500);
        const iters = b.repeatCount === 'indefinite' ? Infinity : (b.repeatCount ? b.repeatCount / 1000 : 1);
        const rev = b.autoRev === '1';
        const opts = { duration: dur, iterations: rev ? iters * 2 : iters, direction: rev ? 'alternate' : 'normal', fill: 'forwards', easing: 'ease-in-out' };
        if (b.t === 'animRot') {
          const by = b.by != null ? b.by : (b.to || 0) - (b.from || 0);
          later(() => track(inn.animate([{ transform: base }, { transform: `${base} rotate(${by}deg)` }], opts)), d);
          end = Math.max(end, d + dur * (rev ? 2 : 1));
        } else if (b.t === 'animScale') {
          const by = b.by || b.to || [1, 1];
          later(() => track(inn.animate([{ transform: base }, { transform: `${base} scale(${by[0]},${by[1]})` }], opts)), d);
          end = Math.max(end, d + dur * (rev ? 2 : 1));
        } else if (b.t === 'animMotion') {
          const pts = parsePath(b.path);
          const fr = pts.map(([x, y], i) => ({ offset: i / (pts.length - 1 || 1), transform: `translate(${x * W}px,${y * H}px)` }));
          later(() => track(target.animate(fr, Object.assign({}, opts, { easing: 'linear', composite: 'replace' }))), d);
          end = Math.max(end, d + (isFinite(iters) ? dur * iters * (rev ? 2 : 1) : dur));
        } else if (b.t === 'cmd') {
          const m = audios[id];
          later(() => {
            if (!m) return;
            const c = b.cmd || '';
            if (/^playFrom/.test(c)) { try { m.currentTime = parseFloat(c.slice(9)) || 0; } catch (_) { } m.play().catch(() => { }); }
            else if (c === 'togglePause') { m.paused ? m.play().catch(() => { }) : m.pause(); }
            else if (c === 'stop' || c === 'pause') m.pause();
            else if (c === 'play' || c === 'resume') m.play().catch(() => { });
          }, d);
        }
      }
    }
    return end;
  }
  function parsePath(p) {
    const tok = (p || '').match(/[MLCZEmlcze]|-?\d*\.?\d+(?:e-?\d+)?/g) || [];
    const pts = []; let cmd = null, buf = [];
    const flush = () => {
      if (!cmd) return;
      const n = cmd.toUpperCase() === 'C' ? 6 : 2;
      for (let i = 0; i + n <= buf.length; i += n) pts.push([+buf[i + n - 2], +buf[i + n - 1]]);
      buf = [];
    };
    for (const t of tok) { if (/[A-Za-z]/.test(t)) { flush(); cmd = t; } else buf.push(t); }
    flush();
    return pts.length ? pts : [[0, 0], [0, 0]];
  }

  function runGroup(effs, instant, done) {
    let end = 0;
    for (const e of effs) {
      if (instant) { runEffect(e, true); continue; }
      const st = e.start;
      later(() => runEffect(e, false), st);
      end = Math.max(end, st + estimate(e));
    }
    if (done) { if (instant) done(); else later(done, end); }
  }
  function estimate(e) {
    let m = 0;
    for (const b of e.b) {
      const it = b.repeatCount === 'indefinite' || (b.repeatCount && +b.repeatCount > 5000) ? 1 : (b.repeatCount ? b.repeatCount / 1000 : 1);
      m = Math.max(m, num(b.delay, 0) + num(b.dur, 0) * it * (b.autoRev === '1' ? 2 : 1));
    }
    return m;
  }

  /* ---------------- main sequence */
  let mainBusy = false;
  function playMainFrom(i, instant) {
    const g = D.slides[cur].tm.main;
    if (i >= g.length) return;
    step = i + 1;
    runGroup(g[i].e, instant, () => {
      if (step < g.length && g[step].auto) playMainFrom(step, instant);
    });
  }
  function next() {
    const g = D.slides[cur].tm.main;
    if (step < g.length) playMainFrom(step, false);
    else go(cur + 1);
  }
  function prev() { if (cur > 0) go(cur - 1, true); }

  /* ---------------- slides */
  let stageInner = null;
  function clearSlide() {
    timers.forEach(clearTimeout); timers = [];
    anims.forEach(a => { try { a.cancel(); } catch (_) { } }); anims = [];
    Object.values(audios).forEach(a => { try { a.pause(); } catch (_) { } });
    audios = {}; els = {}; shapesById = {}; trigState = {};
  }
  function go(i, built) {
    if (i < 0 || i >= D.slides.length) return;
    const old = stageInner;
    clearSlide();
    cur = i; step = 0;
    const sl = D.slides[i];
    stageInner = el('div', 'slide', stage);
    if (sl.bg) stageInner.style.background = sl.bg;
    for (const s of sl.shapes) render(s, stageInner);
    for (const id of sl.hidden) show(id, false);
    // triggers
    for (const id in sl.tm.trig) {
      const t = els[id]; if (!t || t.dataset.media) continue;
      t.classList.add('clickable');
      t.addEventListener('click', ev => {
        ev.stopPropagation();
        const gs = sl.tm.trig[id]; const k = trigState[id] || 0;
        runGroup(gs[k % gs.length], false); trigState[id] = k + 1;
      });
    }
    // transition
    stageInner.animate([{ opacity: 0 }, { opacity: 1 }], { duration: sl.tr ? 450 : 1, easing: 'ease-out' });
    if (old) { old.style.pointerEvents = 'none'; setTimeout(() => old.remove(), 460); }
    const g = sl.tm.main;
    if (built) { for (let k = 0; k < g.length; k++) runGroup(g[k].e, true); step = g.length; for (let k = 0; k < g.length; k++) for (const e of g[k].e) if (e.cls === 'path' || e.cls === 'emph') runEffect(e, false); }
    else if (g.length && g[0].auto) playMainFrom(0, false);
    counter.textContent = `${i + 1} / ${D.slides.length}`;
    document.querySelectorAll('#list li').forEach((li, k) => li.classList.toggle('on', k === i));
    history.replaceState(null, '', '#' + (i + 1));
  }

  /* ---------------- controls */
  stage.addEventListener('click', e => { if (!e.target.closest('.clickable,video,a')) next(); });
  document.addEventListener('keydown', e => {
    if (e.target.closest('input')) return;
    const k = e.key;
    if (['ArrowRight', 'ArrowDown', ' ', 'PageDown', 'Enter', 'n'].includes(k)) { e.preventDefault(); next(); }
    else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p'].includes(k)) { e.preventDefault(); prev(); }
    else if (k === 'Home') go(0); else if (k === 'End') go(D.slides.length - 1);
    else if (k === 'f' || k === 'F') toggleFs();
    else if (k === 'Escape') document.getElementById('menu').classList.remove('open');
  });
  let tx0 = null;
  stage.addEventListener('touchstart', e => { tx0 = e.touches[0].clientX; }, { passive: true });
  stage.addEventListener('touchend', e => {
    if (tx0 == null) return; const dx = e.changedTouches[0].clientX - tx0; tx0 = null;
    if (Math.abs(dx) > 60) { e.preventDefault(); dx < 0 ? next() : prev(); }
  });
  function toggleFs() {
    try { if (!document.fullscreenElement) { const r = (document.documentElement.requestFullscreen || document.documentElement.webkitRequestFullscreen).call(document.documentElement); if (r && r.catch) r.catch(() => { }); } else document.exitFullscreen(); } catch (_) { }
  }
  document.getElementById('bNext').onclick = e => { e.stopPropagation(); next(); };
  document.getElementById('bPrev').onclick = e => { e.stopPropagation(); prev(); };
  document.getElementById('bFs').onclick = e => { e.stopPropagation(); toggleFs(); };
  document.getElementById('bMenu').onclick = e => { e.stopPropagation(); document.getElementById('menu').classList.toggle('open'); };
  const list = document.getElementById('list');
  D.slides.forEach((s, i) => {
    const li = el('li', null, list);
    li.innerHTML = `<b>${i + 1}</b><span></span>`; li.querySelector('span').textContent = s.title || '—';
    li.onclick = () => { document.getElementById('menu').classList.remove('open'); go(i); };
  });
  let hideT;
  function poke() { document.body.classList.add('ui'); clearTimeout(hideT); hideT = setTimeout(() => document.body.classList.remove('ui'), 2500); }
  document.addEventListener('mousemove', poke); document.addEventListener('touchstart', poke, { passive: true });

  /* ---------------- start */
  fit();
  const start = parseInt((location.hash || '').slice(1)) - 1;
  const begin = () => go(isNaN(start) ? 0 : Math.max(0, Math.min(D.slides.length - 1, start)));
  (document.fonts && document.fonts.ready ? document.fonts.ready : Promise.resolve()).then(begin);
  poke();
})();
