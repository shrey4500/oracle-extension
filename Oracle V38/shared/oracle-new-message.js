// oracle-new-message.js — New Message composer slider (Slack/Gmail)
// Supports: 'fullscreen' (newtab) and 'inline' (sidepanel) modes

(function () {
  'use strict';

  const { escapeHtml, WEBHOOK_URL, state } = window.Oracle;

  // Inject new-message CSS (shared across newtab and sidepanel)
  if (!document.getElementById('oracle-new-msg-styles')) {
    const styleEl = document.createElement('style');
    styleEl.id = 'oracle-new-msg-styles';
    styleEl.textContent = `
      .new-msg-recipient-chip { display:inline-flex;align-items:center;gap:6px;padding:4px 10px;background:linear-gradient(45deg,#667eea,#764ba2);color:white;border-radius:16px;font-size:12px;font-weight:500;margin:2px 4px 2px 0; }
      .new-msg-recipient-chip .chip-remove { cursor:pointer;opacity:0.8;font-size:14px;line-height:1; }
      .new-msg-recipient-chip .chip-remove:hover { opacity:1; }
      .new-msg-recipient-chip.channel { background:linear-gradient(45deg,#27ae60,#2ecc71); }
      .new-msg-recipient-chip.group-dm { background:linear-gradient(45deg,#e67e22,#f39c12); }
      .new-msg-search-results { position:absolute;top:100%;left:0;right:0;max-height:200px;overflow-y:auto;border-radius:8px;box-shadow:0 4px 12px rgba(0,0,0,0.15);z-index:100;margin-top:4px; }
      .new-msg-search-item { display:flex;align-items:center;gap:10px;padding:10px 12px;cursor:pointer;transition:background 0.2s; }
      .new-msg-search-item:last-child { border-bottom:none; }
      .new-msg-search-item:hover { background:rgba(102,126,234,0.08); }
      .new-msg-search-item.active { background:rgba(102,126,234,0.12); }
      .new-msg-search-item .item-icon { width:32px;height:32px;border-radius:50%;background:linear-gradient(45deg,#667eea,#764ba2);display:flex;align-items:center;justify-content:center;color:white;font-size:14px;font-weight:600;flex-shrink:0; }
      .new-msg-search-item .item-icon.channel { background:linear-gradient(45deg,#27ae60,#2ecc71);border-radius:8px; }
      .new-msg-search-item .item-icon.group-dm { background:linear-gradient(45deg,#e67e22,#f39c12);border-radius:8px; }
      .new-msg-search-item .item-details { flex:1;min-width:0; }
      .new-msg-search-item .item-name { font-weight:600;font-size:13px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap; }
      .new-msg-search-item .item-email { font-size:11px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;opacity:0.7; }
      .new-msg-type-badge { font-size:10px;padding:2px 8px;border-radius:4px;font-weight:600;flex-shrink:0; }
      .new-msg-type-badge.employee { background:rgba(102,126,234,0.15);color:#667eea; }
      .new-msg-type-badge.channel { background:rgba(39,174,96,0.15);color:#27ae60; }
      .new-msg-type-badge.group-dm { background:rgba(230,126,34,0.15);color:#e67e22; }
      /* Light mode */
      .new-msg-search-results { background:#ffffff;border:1px solid #e1e8ed; }
      .new-msg-search-item { border-bottom:1px solid rgba(225,232,237,0.5); }
      .new-msg-search-item .item-name { color:#2c3e50; }
      .new-msg-search-item .item-email { color:#7f8c8d; }
      /* Dark mode */
      body.dark-mode .new-msg-search-results { background:#1a2332;border:1px solid rgba(255,255,255,0.1); }
      body.dark-mode .new-msg-search-item { border-bottom:1px solid rgba(255,255,255,0.06); }
      body.dark-mode .new-msg-search-item:hover { background:rgba(102,126,234,0.15); }
      body.dark-mode .new-msg-search-item .item-name { color:#e8e8e8; }
      body.dark-mode .new-msg-search-item .item-email { color:#888; }
    `;
    document.head.appendChild(styleEl);
  }

  function showNewMessageSlider(opts = {}) {
    const mode = opts.mode || 'fullscreen';
    const container = opts.container || document.body;
    const onClose = opts.onClose || null;
    const source = opts.source || (mode === 'inline' ? 'oracle-sidepanel' : 'oracle-chrome-extension-newtab');
    const isDark = document.body.classList.contains('dark-mode');

    // Remove existing
    if (mode === 'fullscreen') document.querySelectorAll('.new-message-slider-overlay').forEach(s => s.remove());
    else container.querySelectorAll('.slider-overlay,.new-message-slider-overlay').forEach(s => s.remove());

    const ib = isDark ? 'rgba(255,255,255,0.05)' : 'rgba(102,126,234,0.05)';
    const ibr = isDark ? 'rgba(255,255,255,0.1)' : 'rgba(102,126,234,0.2)';
    const tc = isDark ? '#e8e8e8' : '#2c3e50';
    const lc = isDark ? '#b0b0b0' : '#5d6d7e';
    const bc = isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.5)';

    const formHtml = `
      <div style="padding:20px;border-bottom:1px solid ${bc};display:flex;align-items:center;justify-content:space-between;flex-shrink:0;">
        <div style="display:flex;align-items:center;gap:12px;">
          <div style="width:40px;height:40px;background:linear-gradient(45deg,#667eea,#764ba2);border-radius:10px;display:flex;align-items:center;justify-content:center;font-size:20px;">✉️</div>
          <div><div style="font-weight:600;font-size:16px;color:${tc};">New Message</div><div style="font-size:12px;color:${isDark?'#888':'#7f8c8d'};">Send via Slack or Gmail</div></div>
        </div>
        <button class="new-msg-close-btn" style="background:rgba(231,76,60,0.1);border:none;color:#e74c3c;width:36px;height:36px;border-radius:8px;cursor:pointer;font-size:18px;">×</button>
      </div>
      <div style="flex:1;overflow-y:auto;padding:20px;display:flex;flex-direction:column;gap:16px;">
        <div style="display:flex;flex-direction:column;gap:6px;">
          <label style="font-size:12px;font-weight:600;color:${lc};">Platform</label>
          <select class="nm-platform" style="padding:12px 16px;background:${ib};border:1px solid ${ibr};border-radius:10px;font-size:14px;color:${tc};outline:none;cursor:pointer;"><option value="slack">Slack</option><option value="gmail">Gmail</option></select>
        </div>
        <div style="display:flex;flex-direction:column;gap:6px;position:relative;">
          <label style="font-size:12px;font-weight:600;color:${lc};">To</label>
          <div class="nm-to-wrap" style="display:flex;flex-wrap:wrap;align-items:center;padding:8px 12px;background:${ib};border:1px solid ${ibr};border-radius:10px;min-height:44px;gap:4px;position:relative;">
            <div class="nm-to-chips" style="display:flex;flex-wrap:wrap;gap:4px;"></div>
            <input type="text" class="nm-to-input" placeholder="Search users, channels..." style="flex:1;min-width:120px;border:none;background:transparent;outline:none;font-size:14px;color:${tc};padding:4px 0;">
          </div>
          <div class="nm-to-results new-msg-search-results" style="display:none;"></div>
        </div>
        <div class="nm-cc-section" style="display:none;flex-direction:column;gap:6px;position:relative;">
          <label style="font-size:12px;font-weight:600;color:${lc};">CC</label>
          <div style="display:flex;flex-wrap:wrap;align-items:center;padding:8px 12px;background:${ib};border:1px solid ${ibr};border-radius:10px;min-height:44px;gap:4px;position:relative;">
            <div class="nm-cc-chips" style="display:flex;flex-wrap:wrap;gap:4px;"></div>
            <input type="text" class="nm-cc-input" placeholder="Search or type email..." style="flex:1;min-width:120px;border:none;background:transparent;outline:none;font-size:14px;color:${tc};padding:4px 0;">
          </div>
          <div class="nm-cc-results new-msg-search-results" style="display:none;"></div>
        </div>
        <div class="nm-subject-section" style="display:none;flex-direction:column;gap:6px;">
          <label style="font-size:12px;font-weight:600;color:${lc};">Subject</label>
          <input type="text" class="nm-subject" placeholder="Email subject..." style="padding:12px 16px;background:${ib};border:1px solid ${ibr};border-radius:10px;font-size:14px;color:${tc};outline:none;">
        </div>
        <div style="display:flex;flex-direction:column;gap:6px;flex:1;position:relative;">
          <label style="font-size:12px;font-weight:600;color:${lc};">Message</label>
          <textarea class="nm-body" placeholder="Type your message... Use @name to mention someone" style="flex:1;min-height:150px;padding:12px 16px;background:${ib};border:1px solid ${ibr};border-radius:10px;font-size:14px;color:${tc};outline:none;resize:none;font-family:inherit;line-height:1.5;"></textarea>
          <div class="nm-mention-dropdown new-msg-search-results" style="display:none;bottom:100%;top:auto;margin-bottom:4px;max-height:180px;"></div>
        </div>
      </div>
      <div style="padding:16px 20px;border-top:1px solid ${bc};flex-shrink:0;">
        <button class="nm-send-btn" style="width:100%;padding:14px;background:linear-gradient(45deg,#667eea,#764ba2);border:none;color:white;border-radius:10px;font-size:15px;font-weight:600;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:8px;"><span>Send Message</span><span>➤</span></button>
      </div>`;

    let slider, overlay, closeSlider;

    if (mode === 'inline') {
      slider = document.createElement('div');
      slider.className = 'slider-overlay';
      slider.innerHTML = formHtml;
      container.appendChild(slider);
      closeSlider = () => { slider.classList.add('closing'); if (onClose) onClose(); setTimeout(() => slider.remove(), 250); };
      slider.querySelector('.new-msg-close-btn').addEventListener('click', closeSlider);
    } else {
      overlay = document.createElement('div');
      overlay.className = 'new-message-slider-overlay';
      // Position overlay to cover col3 only (like transcript slider)
      const col3Rect = window.Oracle.getCol3Rect();
      if (col3Rect) {
        overlay.style.cssText = `position:fixed;top:${col3Rect.top}px;left:${col3Rect.left}px;width:${col3Rect.width}px;bottom:0;z-index:10000;display:flex;flex-direction:column;animation:fadeIn 0.15s ease-out;`;
      } else {
        overlay.style.cssText = 'position:fixed;top:0;right:0;bottom:0;width:400px;z-index:10000;display:flex;flex-direction:column;animation:fadeIn 0.15s ease-out;';
      }
      slider = document.createElement('div');
      slider.className = 'new-message-slider';
      slider.style.cssText = `width:100%;height:100%;background:${isDark?'#1f2940':'white'};box-shadow:-4px 0 20px rgba(0,0,0,${isDark?'0.4':'0.15'});display:flex;flex-direction:column;animation:slideInRight 0.3s ease-out;border-radius:12px;border:1px solid ${isDark?'rgba(255,255,255,0.08)':'rgba(225,232,237,0.6)'};overflow:hidden;`;
      slider.innerHTML = formHtml;
      overlay.appendChild(slider);
      document.body.appendChild(overlay);
      closeSlider = () => { overlay.style.animation='fadeOut 0.2s ease-out'; slider.style.animation='slideOutRight 0.3s ease-out'; document.removeEventListener('keydown', kh); if (onClose) onClose(); setTimeout(()=>{ overlay.remove(); window.Oracle.collapseCol3AfterSlider(); },250); };
      slider.querySelector('.new-msg-close-btn').addEventListener('click', closeSlider);
      overlay.addEventListener('click', (e) => { if (e.target===overlay) closeSlider(); });
      const kh = (e) => { if (e.key==='Escape') closeSlider(); };
      document.addEventListener('keydown', kh);
    }

    // --- Wire up logic ---
    let platform = 'slack', toR = [], ccR = [], sTO = null, hasCh = false;
    const mentionsMap = new Map();
    let mentionStart = -1, mentionTO = null, mentionActive = false;
    const q = (s) => slider.querySelector(s);
    const plat = q('.nm-platform'), toIn = q('.nm-to-input'), toRes = q('.nm-to-results'), toChips = q('.nm-to-chips');
    const ccSec = q('.nm-cc-section'), ccIn = q('.nm-cc-input'), ccRes = q('.nm-cc-results'), ccChips = q('.nm-cc-chips');
    const subSec = q('.nm-subject-section'), subIn = q('.nm-subject'), body = q('.nm-body'), sendBtn = q('.nm-send-btn'), mDD = q('.nm-mention-dropdown');

    plat.addEventListener('change', () => {
      platform = plat.value; toR=[]; ccR=[]; hasCh=false; rTC(); rCC(); toIn.value=''; ccIn.value='';
      toIn.placeholder = platform==='slack' ? 'Search users, channels...' : 'Search or type email...';
      toIn.disabled = false;
      ccSec.style.display = platform==='gmail' ? 'flex' : 'none';
      subSec.style.display = platform==='gmail' ? 'flex' : 'none';
    });

    async function search(query, target, isCC) {
      if (query.length < 3) { target.style.display='none'; return; }
      const controller = new AbortController();
      const tid = setTimeout(() => controller.abort(), 15000);
      try {
        const wn = platform==='slack' ? 'search_user_new_slack' : 'search_user_new_gmail';
        const eT = toR.map(r=>r.id).filter(Boolean), eC = ccR.map(r=>r.id).filter(Boolean);
        const res = await fetch(WEBHOOK_URL, { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({ action:wn, query, platform, existing_to_ids:eT, existing_cc_ids:isCC?eC:[], exclude_ids:isCC?[...eT,...eC]:eT, timestamp:new Date().toISOString(), source, user_id:state.userData?.userId, authenticated:true }), signal: controller.signal });
        clearTimeout(tid);
        if (res.ok) { let d=await res.json(); if(!Array.isArray(d)) d=d.results||d.members||d.users||[]; renderRes(d,target,isCC); }
      } catch(e) { clearTimeout(tid); console.error('Search error:',e); target.style.display='none'; }
    }

    function renderRes(results, target, isCC) {
      if (!results?.length||(platform==='slack'&&hasCh&&!isCC)) { target.style.display='none'; return; }
      const sids = (isCC?ccR:toR).map(r=>r.id||r.user_slack_ID);
      const f = results.filter(r=>!sids.includes(r.id||r.user_slack_ID));
      if (!f.length) { target.style.display='none'; return; }
      target.innerHTML = f.map(item => {
        let t=item.type||'employee', dn=item.Full_Name||item['Full Name']||item.full_name||item.name||item.user_email_ID||'Unknown', sub=item.user_email_ID||item.email||'';
        if (t==='Private Channel'||t==='channel') { t='channel'; sub='Channel'; }
        else if (t==='group_dm'||t==='mpim') { t='group_dm'; dn=item.Full_Name||item.name||'Group DM'; sub='Group DM'; }
        const ini=dn.split(' ').map(n=>n[0]).join('').substring(0,2).toUpperCase();
        return `<div class="new-msg-search-item" data-id="${item.user_slack_ID||item.slack_id||item.id||''}" data-name="${escapeHtml(dn)}" data-email="${escapeHtml(item.user_email_ID||item.email||'')}" data-type="${t}" data-is-cc="${isCC}"><div class="item-icon ${t==='channel'?'channel':t==='group_dm'?'group-dm':''}">${t==='channel'?'#':t==='group_dm'?'👥':ini}</div><div class="item-details"><div class="item-name">${escapeHtml(dn)}</div><div class="item-email">${escapeHtml(sub)}</div></div><span class="new-msg-type-badge ${t==='channel'?'channel':t==='group_dm'?'group-dm':'employee'}">${t==='channel'?'Channel':t==='group_dm'?'Group DM':'Employee'}</span></div>`;
      }).join('');
      target.style.display='block';
      target.querySelectorAll('.new-msg-search-item').forEach(it => {
        it.addEventListener('click', () => {
          const r={id:it.dataset.id, name:it.dataset.name, email:it.dataset.email, type:it.dataset.type};
          if (it.dataset.isCc==='true') { ccR.push(r); rCC(); ccIn.value=''; ccRes.style.display='none'; }
          else addTo(r);
          toIn.focus();
        });
      });
    }

    function addTo(r) {
      if (platform==='slack'&&(r.type==='channel'||r.type==='group_dm')) { toR=[r]; hasCh=true; toIn.placeholder='Channel/Group selected'; toIn.disabled=true; }
      else if (platform==='slack'&&hasCh) return;
      else toR.push(r);
      rTC(); toIn.value=''; toRes.style.display='none';
    }

    function rTC() {
      toChips.innerHTML=toR.map((r,i)=>`<span class="new-msg-recipient-chip ${r.type==='channel'?'channel':r.type==='group_dm'?'group-dm':''}" data-idx="${i}">${r.type==='channel'?'#':r.type==='group_dm'?'👥':''}${escapeHtml(r.name||r.email)}<span class="chip-remove" data-idx="${i}">×</span></span>`).join('');
      toChips.querySelectorAll('.chip-remove').forEach(b=>{ b.addEventListener('click',(e)=>{ e.stopPropagation(); const rem=toR.splice(parseInt(b.dataset.idx),1)[0]; if(rem&&(rem.type==='channel'||rem.type==='group_dm')){hasCh=false;toIn.placeholder='Search users, channels...';toIn.disabled=false;} rTC(); }); });
    }
    function rCC() {
      ccChips.innerHTML=ccR.map((r,i)=>`<span class="new-msg-recipient-chip" data-idx="${i}">${escapeHtml(r.name||r.email)}<span class="chip-remove" data-idx="${i}">×</span></span>`).join('');
      ccChips.querySelectorAll('.chip-remove').forEach(b=>{ b.addEventListener('click',(e)=>{ e.stopPropagation(); ccR.splice(parseInt(b.dataset.idx),1); rCC(); }); });
    }

    function setupNav(inp, res, isCC) {
      inp.addEventListener('input', ()=>{ clearTimeout(sTO); sTO=setTimeout(()=>search(inp.value.trim(),res,isCC),300); });
      inp.addEventListener('keydown', (e)=>{
        const items=res.querySelectorAll('.new-msg-search-item'), act=res.querySelector('.new-msg-search-item.active');
        if (res.style.display!=='none'&&items.length>0) {
          let idx=Array.from(items).indexOf(act);
          if (e.key==='ArrowDown') { e.preventDefault(); act?.classList.remove('active'); idx=(idx+1)%items.length; items[idx].classList.add('active'); items[idx].scrollIntoView({block:'nearest'}); return; }
          if (e.key==='ArrowUp') { e.preventDefault(); act?.classList.remove('active'); idx=idx<=0?items.length-1:idx-1; items[idx].classList.add('active'); items[idx].scrollIntoView({block:'nearest'}); return; }
          if (e.key==='Enter') { e.preventDefault(); (act||(items.length===1?items[0]:null))?.click(); return; }
          if (e.key==='Escape') { res.style.display='none'; return; }
        }
        if (e.key==='Enter'&&platform==='gmail') {
          e.preventDefault(); const em=inp.value.trim();
          if (em&&em.includes('@')) { const r={id:em,name:em,email:em,type:'email'}; if(isCC){ccR.push(r);rCC();inp.value='';ccRes.style.display='none';}else addTo(r); }
        }
      });
    }
    setupNav(toIn, toRes, false);
    setupNav(ccIn, ccRes, true);

    // @mention in body
    body.addEventListener('input', ()=>{
      const txt=body.value, cp=body.selectionStart, before=txt.substring(0,cp), mm=before.match(/@(\w*)$/);
      if (mm) { const q=mm[1]; mentionStart=cp-q.length-1; if(q.length>=3){mentionActive=true;clearTimeout(mentionTO);mentionTO=setTimeout(()=>searchMention(q),300);}else{mDD.style.display='none';mentionActive=false;}}
      else { mDD.style.display='none'; mentionStart=-1; mentionActive=false; }
    });

    async function searchMention(query) {
      const controller = new AbortController();
      const tid = setTimeout(() => controller.abort(), 15000);
      try {
        const wn=platform==='slack'?'search_user_new_slack':'search_user_new_gmail';
        const res=await fetch(WEBHOOK_URL,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({action:wn,query,platform,context:'mention',timestamp:new Date().toISOString(),source,user_id:state.userData?.userId,authenticated:true}),signal:controller.signal});
        clearTimeout(tid);
        if (res.ok) { let d=await res.json(); if(!Array.isArray(d)) d=d.results||d.members||d.users||[]; renderMention(d.filter(r=>r.type==='employee'||!r.type).slice(0,5)); }
      } catch(e) { clearTimeout(tid); console.error('Mention search error:',e); mDD.style.display='none'; }
    }

    function renderMention(employees) {
      if (!employees?.length) { mDD.style.display='none'; return; }
      mDD.innerHTML=employees.map(item=>{
        const dn=item.Full_Name||item['Full Name']||item.full_name||item.name||'Unknown';
        const ini=dn.split(' ').map(n=>n[0]).join('').substring(0,2).toUpperCase();
        return `<div class="new-msg-search-item mention-item" data-id="${item.user_slack_ID||item.slack_id||item.id||''}" data-name="${escapeHtml(dn)}" data-email="${escapeHtml(item.user_email_ID||item.email||'')}"><div class="item-icon">${ini}</div><div class="item-details"><div class="item-name">${escapeHtml(dn)}</div><div class="item-email">${escapeHtml(item.user_email_ID||item.email||'')}</div></div></div>`;
      }).join('');
      mDD.style.display='block';
      mDD.querySelectorAll('.mention-item').forEach(it=>{ it.addEventListener('click',()=>insertMention(it.dataset.id,it.dataset.name)); });
    }

    function insertMention(uid, uname) {
      const txt=body.value, cp=body.selectionStart, before=txt.substring(0,cp), mm=before.match(/@(\w*)$/);
      if (mm&&mentionStart>=0) {
        const dt=`@${uname}`; mentionsMap.set(dt,{id:uid,name:uname});
        body.value=txt.substring(0,mentionStart)+dt+' '+txt.substring(cp);
        const np=mentionStart+dt.length+1; body.setSelectionRange(np,np);
      }
      mDD.style.display='none'; mentionStart=-1; mentionActive=false; body.focus();
    }

    body.addEventListener('keydown', (e)=>{
      if (!mentionActive||mDD.style.display==='none') return;
      const items=mDD.querySelectorAll('.mention-item'), act=mDD.querySelector('.mention-item.active');
      let idx=Array.from(items).indexOf(act);
      if (e.key==='ArrowDown'){e.preventDefault();act?.classList.remove('active');idx=(idx+1)%items.length;items[idx].classList.add('active');}
      else if(e.key==='ArrowUp'){e.preventDefault();act?.classList.remove('active');idx=idx<=0?items.length-1:idx-1;items[idx].classList.add('active');}
      else if(e.key==='Enter'&&act){e.preventDefault();insertMention(act.dataset.id,act.dataset.name);}
      else if(e.key==='Escape'){e.preventDefault();mDD.style.display='none';mentionActive=false;}
      else if(e.key==='Tab'&&items.length>0){e.preventDefault();const f=act||items[0];insertMention(f.dataset.id,f.dataset.name);}
    });

    // Send
    sendBtn.addEventListener('click', async ()=>{
      if (!toR.length) { alert('Please add at least one recipient.'); return; }
      let msg=body.value.trim(); if(!msg){alert('Please enter a message.');return;}
      // Convert mentions
      mentionsMap.forEach((d,dt)=>{ if(platform==='slack') msg=msg.replace(new RegExp(dt.replace(/[.*+?^${}()|[\]\\]/g,'\\$&'),'g'),`<@${d.id}>`); });
      sendBtn.disabled=true; sendBtn.innerHTML='<span>Sending...</span>';
      try {
        const wn=platform==='slack'?'send_new_slack_message':'send_new_gmail_email';
        const payload={action:wn,platform,to:toR.map(r=>({id:r.id,name:r.name,email:r.email,type:r.type})),message:msg,timestamp:new Date().toISOString(),source,user_id:state.userData?.userId,authenticated:true};
        if (platform==='gmail') { payload.cc=ccR.map(r=>({id:r.id,name:r.name,email:r.email,type:r.type})); payload.subject=subIn.value.trim()||'(No Subject)'; }
        const res=await fetch(WEBHOOK_URL,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
        if (res.ok) { if(typeof showToastNotification==='function') showToastNotification(`Message sent via ${platform==='slack'?'Slack':'Gmail'}!`); closeSlider(); }
        else throw new Error('Failed to send message');
      } catch(e) { console.error('Send error:',e); alert('Failed to send: '+e.message); sendBtn.disabled=false; sendBtn.innerHTML='<span>Send Message</span><span>➤</span>'; }
    });

    setTimeout(()=>toIn.focus(),100);
  }

  window.OracleNewMessage = { showNewMessageSlider };
})();
