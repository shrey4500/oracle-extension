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
      .new-msg-search-item .item-icon.private-channel { background:linear-gradient(45deg,#8e44ad,#9b59b6);border-radius:8px; }
      .new-msg-search-item .item-icon.group-dm { background:linear-gradient(45deg,#e67e22,#f39c12);border-radius:8px; }
      .new-msg-search-item .item-details { flex:1;min-width:0; }
      .new-msg-search-item .item-name { font-weight:600;font-size:13px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap; }
      .new-msg-search-item .item-name.wrap { white-space:normal;display:-webkit-box;-webkit-line-clamp:2;-webkit-box-orient:vertical; }
      .new-msg-search-item .item-email { font-size:11px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;opacity:0.7; }
      .new-msg-type-badge { font-size:10px;padding:2px 8px;border-radius:4px;font-weight:600;flex-shrink:0; }
      .new-msg-type-badge.dm { background:rgba(102,126,234,0.15);color:#667eea; }
      .new-msg-type-badge.employee { background:rgba(102,126,234,0.15);color:#667eea; }
      .new-msg-type-badge.channel { background:rgba(39,174,96,0.15);color:#27ae60; }
      .new-msg-type-badge.private-channel { background:rgba(142,68,173,0.15);color:#8e44ad; }
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
      .nm-body:empty:before { content:attr(data-placeholder);color:#999;pointer-events:none; }
      .nm-body a { color:#667eea;text-decoration:underline;cursor:pointer; }
      body.dark-mode .nm-body a { color:#8fa4f8; }
      .nm-body b, .nm-body strong { font-weight:700; }
      .nm-body i, .nm-body em { font-style:italic; }
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
          <div class="nm-platform-toggle" style="display:flex;gap:8px;">
            <button class="nm-platform-btn nm-platform-btn-active" data-platform="slack" style="flex:1;padding:10px 0;border-radius:10px;border:2px solid #667eea;background:rgba(102,126,234,0.12);color:#667eea;font-size:14px;font-weight:600;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:6px;transition:all 0.2s;"><img src="icon-slack.png" width="18" height="18" style="object-fit:contain;flex-shrink:0;"> Slack</button>
            <button class="nm-platform-btn" data-platform="gmail" style="flex:1;padding:10px 0;border-radius:10px;border:2px solid ${ibr};background:${ib};color:${tc};font-size:14px;font-weight:600;cursor:pointer;display:flex;align-items:center;justify-content:center;gap:6px;transition:all 0.2s;"><img src="icon-gmail.png" width="18" height="18" style="object-fit:contain;flex-shrink:0;"> Gmail</button>
          </div>
          <select class="nm-platform" style="display:none;"><option value="slack">Slack</option><option value="gmail">Gmail</option></select>
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
        <div style="display:flex;flex-direction:column;gap:6px;">
          <label style="font-size:12px;font-weight:600;color:${lc};">Attachments</label>
          <div style="display:flex;align-items:center;gap:8px;">
            <input type="file" class="nm-file-input" style="display:none;" multiple>
            <button class="nm-attach-btn" style="padding:8px 14px;background:${ib};border:1px solid ${ibr};border-radius:10px;font-size:13px;color:${tc};cursor:pointer;display:flex;align-items:center;gap:6px;">📎 Attach Files</button>
          </div>
          <div class="nm-attach-chips" style="display:flex;flex-wrap:wrap;gap:6px;"></div>
        </div>
        <div style="display:flex;flex-direction:column;gap:6px;flex:1;position:relative;">
          <label style="font-size:12px;font-weight:600;color:${lc};">Message</label>
          <div class="nm-body" contenteditable="true" data-placeholder="Type your message... Use @name to mention someone" style="flex:1;min-height:150px;padding:12px 16px;background:${ib};border:1px solid ${ibr};border-radius:10px;font-size:14px;color:${tc};outline:none;resize:none;font-family:inherit;line-height:1.5;overflow-y:auto;white-space:pre-wrap;word-wrap:break-word;"></div>
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
    } else if (mode === 'col3') {
      // Extension: overlay exactly over col3 column, like transcript slider
      const col3Rect = window.Oracle?.getCol3Rect?.() || null;
      overlay = document.createElement('div');
      overlay.className = 'new-message-slider-overlay';
      if (col3Rect) {
        overlay.style.cssText = `position:fixed;top:${col3Rect.top}px;left:${col3Rect.left}px;width:${col3Rect.width}px;bottom:0;z-index:10000;display:flex;flex-direction:column;animation:fadeIn 0.15s ease-out;`;
      } else {
        overlay.style.cssText = 'position:fixed;top:0;right:0;bottom:0;width:420px;z-index:10000;display:flex;flex-direction:column;animation:fadeIn 0.15s ease-out;';
      }
      slider = document.createElement('div');
      slider.className = 'new-message-slider';
      slider.style.cssText = `width:100%;height:100%;background:${isDark?'#1f2940':'white'};box-shadow:-4px 0 20px rgba(0,0,0,${isDark?'0.4':'0.15'});display:flex;flex-direction:column;animation:slideInRight 0.3s ease-out;border-radius:12px;border:1px solid ${isDark?'rgba(255,255,255,0.08)':'rgba(225,232,237,0.6)'};`;
      slider.innerHTML = formHtml;
      overlay.appendChild(slider);
      document.body.appendChild(overlay);
      const kh = (e) => { if (e.key==='Escape') closeSlider(); };
      closeSlider = () => { overlay.style.animation='fadeOut 0.2s ease-out'; slider.style.animation='slideOutRight 0.3s ease-out'; document.removeEventListener('keydown', kh); if (onClose) onClose(); if (window.Oracle?.collapseCol3AfterSlider) window.Oracle.collapseCol3AfterSlider(); setTimeout(()=>overlay.remove(),250); };
      slider.querySelector('.new-msg-close-btn').addEventListener('click', closeSlider);
      document.addEventListener('keydown', kh);
    } else {
      overlay = document.createElement('div');
      overlay.className = 'new-message-slider-overlay';
      // PWA: position below header, full width
      const mainArea = document.getElementById('mainContentArea');
      const headerEl = document.querySelector('.header');
      if (mainArea && headerEl) {
        const headerH = headerEl.getBoundingClientRect().bottom;
        overlay.style.cssText = `position:fixed;top:${headerH}px;left:0;right:0;bottom:0;z-index:10000;display:flex;flex-direction:column;animation:fadeIn 0.15s ease-out;`;
      } else {
        overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.4);z-index:10000;display:flex;justify-content:flex-end;animation:fadeIn 0.2s ease-out;';
      }
      slider = document.createElement('div');
      slider.className = 'new-message-slider';
      slider.style.cssText = `width:100%;height:100%;background:${isDark?'#1f2940':'white'};box-shadow:-4px 0 20px rgba(0,0,0,${isDark?'0.4':'0.15'});display:flex;flex-direction:column;animation:slideInRight 0.3s ease-out;`;
      slider.innerHTML = formHtml;
      overlay.appendChild(slider);
      document.body.appendChild(overlay);
      closeSlider = () => { overlay.style.animation='fadeOut 0.2s ease-out'; slider.style.animation='slideOutRight 0.3s ease-out'; document.removeEventListener('keydown', kh); if (onClose) onClose(); setTimeout(()=>overlay.remove(),250); };
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

    // --- Attachment handling ---
    const nmPendingAttachments = [];
    const fileInput = q('.nm-file-input');
    const attachBtn = q('.nm-attach-btn');
    const attachChips = q('.nm-attach-chips');

    function fileToBase64(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = reject;
        reader.readAsDataURL(file);
      });
    }

    function renderNmAttachChips() {
      if (!attachChips) return;
      attachChips.innerHTML = nmPendingAttachments.map((a, i) => {
        const sizeStr = a.size > 1048576 ? (a.size / 1048576).toFixed(1) + ' MB' : (a.size / 1024).toFixed(0) + ' KB';
        return `<div style="display:flex;align-items:center;gap:6px;padding:6px 10px;background:${isDark ? 'rgba(255,255,255,0.06)' : 'rgba(102,126,234,0.06)'};border:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.6)'};border-radius:8px;font-size:12px;color:${tc};max-width:200px;">
          <span style="font-size:14px;">${a.type?.startsWith('image/') ? '🖼️' : '📄'}</span>
          <span style="flex:1;min-width:0;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${escapeHtml(a.name)}</span>
          <span style="font-size:10px;color:${isDark ? '#888' : '#95a5a6'};flex-shrink:0;">${sizeStr}</span>
          <span class="nm-attach-remove" data-idx="${i}" style="cursor:pointer;font-size:14px;color:#e74c3c;flex-shrink:0;">×</span>
        </div>`;
      }).join('');
      attachChips.querySelectorAll('.nm-attach-remove').forEach(btn => {
        btn.addEventListener('click', () => { nmPendingAttachments.splice(parseInt(btn.dataset.idx), 1); renderNmAttachChips(); });
      });
    }

    if (attachBtn && fileInput) {
      attachBtn.addEventListener('click', () => fileInput.click());
      fileInput.addEventListener('change', async () => {
        for (const file of fileInput.files) {
          if (file.size > 10 * 1024 * 1024) { alert(`File "${file.name}" exceeds 10MB limit.`); continue; }
          try {
            const data = await fileToBase64(file);
            nmPendingAttachments.push({ name: file.name, type: file.type, size: file.size, data });
          } catch (e) { console.error('File read error:', e); }
        }
        renderNmAttachChips();
        fileInput.value = '';
      });
    }

    plat.addEventListener('change', () => {
      platform = plat.value; toR=[]; ccR=[]; hasCh=false; rTC(); rCC(); toIn.value=''; ccIn.value='';
      toRes.style.display='none'; ccRes.style.display='none';
      toIn.placeholder = platform==='slack' ? 'Search users, channels...' : 'Search or type email...';
      toIn.disabled = false;
      ccSec.style.display = platform==='gmail' ? 'flex' : 'none';
      subSec.style.display = platform==='gmail' ? 'flex' : 'none';
    });

    // Wire toggle buttons — directly update platform + show/hide Gmail-only fields
    const ib_ = ib, ibr_ = ibr, tc_ = tc;
    slider.querySelectorAll('.nm-platform-btn').forEach(btn => {
      btn.addEventListener('click', () => {
        const p = btn.dataset.platform;
        platform = p;
        plat.value = p;
        // Reset recipients
        toR=[]; ccR=[]; hasCh=false; rTC(); rCC(); toIn.value=''; ccIn.value='';
        toRes.style.display='none'; ccRes.style.display='none';
        toIn.placeholder = p==='slack' ? 'Search users, channels...' : 'Search or type email...';
        toIn.disabled = false;
        // Show CC + Subject only for Gmail
        ccSec.style.display = p==='gmail' ? 'flex' : 'none';
        subSec.style.display = p==='gmail' ? 'flex' : 'none';
        // Update button styles
        slider.querySelectorAll('.nm-platform-btn').forEach(b => {
          const isActive = b.dataset.platform === p;
          b.classList.toggle('nm-platform-btn-active', isActive);
          b.style.border = isActive ? '2px solid #667eea' : `2px solid ${ibr_}`;
          b.style.background = isActive ? 'rgba(102,126,234,0.12)' : ib_;
          b.style.color = isActive ? '#667eea' : tc_;
        });
      });
    });

    async function search(query, target, isCC) {
      if (query.length < 3) { target.style.display='none'; return; }
      try {
        const wn = platform==='slack' ? 'search_user_new_slack' : 'search_user_new_gmail';
        const eT = toR.map(r=>r.id).filter(Boolean), eC = ccR.map(r=>r.id).filter(Boolean);
        const controller = new AbortController();
        const tid = setTimeout(() => controller.abort(), 15000);
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
        let isPrivate=false;
        if (t==='Private Channel') { isPrivate=true; t='channel'; sub='Private Channel'; }
        else if (t==='Public Channel'||t==='channel') { t='channel'; sub='Public Channel'; }
        else if (t==='Group DM'||t==='group_dm'||t==='mpim') { t='group_dm'; dn=item.Full_Name||item['Full Name']||item.name||'Group DM'; sub='Group DM'; }
        else if (t==='Direct Message'||t==='employee') { t='employee'; sub=item.user_email_ID||item.email||''; }
        const ini=dn.split(' ').map(n=>n[0]).join('').substring(0,2).toUpperCase();
        const iconClass=isPrivate?'private-channel':t==='channel'?'channel':t==='group_dm'?'group-dm':'';
        const iconContent=isPrivate?'🔒':t==='channel'?'#':t==='group_dm'?'👥':ini;
        const badgeClass=isPrivate?'private-channel':t==='channel'?'channel':t==='group_dm'?'group-dm':'dm';
        const badgeLabel=isPrivate?'Private Channel':t==='channel'?'Public Channel':t==='group_dm'?'Group DM':'Direct Message';
        const badgeHtml=platform==='gmail'?'':`<span class="new-msg-type-badge ${badgeClass}">${badgeLabel}</span>`;
        const nameClass=t==='group_dm'?'item-name wrap':'item-name';
        return `<div class="new-msg-search-item" data-id="${item.user_slack_ID||item.slack_id||item.id||''}" data-name="${escapeHtml(dn)}" data-email="${escapeHtml(item.user_email_ID||item.email||'')}" data-type="${t}" data-is-cc="${isCC}"><div class="item-icon ${iconClass}">${iconContent}</div><div class="item-details"><div class="${nameClass}">${escapeHtml(dn)}</div><div class="item-email">${escapeHtml(sub)}</div></div>${badgeHtml}</div>`;
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
      inp.addEventListener('input', ()=>{ const v=inp.value.trim(); if(!v){res.style.display='none';return;} clearTimeout(sTO); sTO=setTimeout(()=>search(v,res,isCC),300); });
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

    // @mention in body — walk backwards to find @ (allow spaces in names)
    body.addEventListener('input', ()=>{
      const sel = window.getSelection();
      if (!sel.rangeCount || !body.contains(sel.anchorNode)) { mDD.style.display='none'; mentionActive=false; return; }
      // Get text before cursor
      const range = sel.getRangeAt(0);
      const preRange = document.createRange();
      preRange.selectNodeContents(body);
      preRange.setEnd(range.startContainer, range.startOffset);
      const before = preRange.toString();
      let atPos = -1;
      for (let i = before.length - 1; i >= 0; i--) {
        if (before[i] === '@') { atPos = i; break; }
        if (before[i] === '\n') break;
      }
      if (atPos !== -1) {
        const q = before.substring(atPos + 1).trimEnd();
        mentionStart = atPos;
        if(q.length>=2){mentionActive=true;clearTimeout(mentionTO);mentionTO=setTimeout(()=>searchMention(q),300);}else{mDD.style.display='none';mentionActive=false;}
      }
      else { mDD.style.display='none'; mentionStart=-1; mentionActive=false; }
    });

    async function searchMention(query) {
      try {
        const wn=platform==='slack'?'search_user_new_slack':'search_user_new_gmail';
        const controller = new AbortController();
        const tid = setTimeout(() => controller.abort(), 15000);
        const queryWords = query.toLowerCase().split(/\s+/).filter(w => w.length > 0);
        const backendQuery = queryWords.length > 1 ? queryWords[0] : query;
        console.log('🔍 NM mention search:', query, 'backend query:', backendQuery);
        const res=await fetch(WEBHOOK_URL,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({action:wn,query:backendQuery,platform,context:'mention',timestamp:new Date().toISOString(),source,user_id:state.userData?.userId,authenticated:true}),signal:controller.signal});
        clearTimeout(tid);
        if (res.ok) {
          let d=await res.json();
          console.log('🔍 NM mention raw response:', d);
          if(!Array.isArray(d)) d=d.results||d.members||d.users||[];
          const filtered = d.filter(r => {
            const t = (r.type || '').toLowerCase();
            const isPersonLike = !t || t === 'employee' || t === 'direct message' || t === 'dm' || t === 'user' || t === 'member';
            return isPersonLike && (r.user_email_ID || r.email || r.Full_Name || r['Full Name'] || r.full_name || r.name);
          });
          // Client-side multi-word filtering
          const multiWordFiltered = queryWords.length > 1
            ? filtered.filter(r => {
                const name = (r.Full_Name || r['Full Name'] || r.full_name || r.name || '').toLowerCase();
                const email = (r.user_email_ID || r.email || '').toLowerCase();
                const haystack = name + ' ' + email;
                return queryWords.every(w => haystack.includes(w));
              })
            : filtered;
          console.log('🔍 NM mention filtered:', multiWordFiltered.length, '(query words:', queryWords, ')');
          renderMention(multiWordFiltered.slice(0, 5));
        } else { console.warn('🔍 NM mention HTTP error:', res.status); }
      } catch(e) { if(e.name!=='AbortError') console.error('Mention search error:',e); mDD.style.display='none'; }
    }

    function renderMention(employees) {
      if (!employees?.length) { mDD.style.display='none'; return; }
      mDD.innerHTML=employees.map(item=>{
        const dn=item.Full_Name||item['Full Name']||item.full_name||item.name||'Unknown';
        const ini=dn.split(' ').map(n=>n[0]).join('').substring(0,2).toUpperCase();
        return `<div class="new-msg-search-item mention-item" data-id="${item.user_slack_ID||item.slack_id||item.id||''}" data-name="${escapeHtml(dn)}" data-email="${escapeHtml(item.user_email_ID||item.email||'')}"><div class="item-icon">${ini}</div><div class="item-details"><div class="item-name">${escapeHtml(dn)}</div><div class="item-email">${escapeHtml(item.user_email_ID||item.email||'')}</div></div></div>`;
      }).join('');
      mDD.style.display='block';
      // Prevent mousedown from stealing focus from body (contenteditable)
      mDD.querySelectorAll('.mention-item').forEach(it=>{ 
        it.addEventListener('mousedown',(e)=>e.preventDefault());
        it.addEventListener('click',()=>insertMention(it.dataset.id,it.dataset.name)); 
      });
    }

    function insertMention(uid, uname) {
      const sel = window.getSelection();
      if (!sel.rangeCount || mentionStart < 0) { mDD.style.display='none'; mentionActive=false; return; }
      // Find and delete from @ to cursor using text content traversal
      const fullText = body.textContent;
      const range = sel.getRangeAt(0);
      // Walk body to find the text node and offset for mentionStart
      let charCount = 0, startNode = null, startOffset = 0;
      const walker = document.createTreeWalker(body, NodeFilter.SHOW_TEXT, null, false);
      while (walker.nextNode()) {
        const node = walker.currentNode;
        if (charCount + node.textContent.length > mentionStart) {
          startNode = node;
          startOffset = mentionStart - charCount;
          break;
        }
        charCount += node.textContent.length;
      }
      if (startNode) {
        const delRange = document.createRange();
        delRange.setStart(startNode, startOffset);
        delRange.setEnd(range.startContainer, range.startOffset);
        delRange.deleteContents();
        // Insert mention as a styled span
        const mentionSpan = document.createElement('span');
        mentionSpan.className = 'mention-tag';
        mentionSpan.dataset.slackId = uid;
        mentionSpan.dataset.name = uname;
        mentionSpan.contentEditable = 'false';
        mentionSpan.style.cssText = 'display:inline-flex;align-items:center;background:rgba(102,126,234,0.15);color:#667eea;padding:2px 6px;border-radius:6px;font-size:13px;font-weight:600;margin:0 2px;';
        mentionSpan.textContent = `@${uname}`;
        mentionsMap.set(`@${uname}`, {id: uid, name: uname});
        const insertRange = window.getSelection().getRangeAt(0);
        insertRange.insertNode(mentionSpan);
        // Add a space after and place cursor there
        const space = document.createTextNode('\u00A0');
        mentionSpan.parentNode.insertBefore(space, mentionSpan.nextSibling);
        const newRange = document.createRange();
        newRange.setStartAfter(space);
        newRange.collapse(true);
        sel.removeAllRanges();
        sel.addRange(newRange);
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
      const bodyText = body.textContent.trim();
      if(!bodyText){alert('Please enter a message.');return;}

      let msg;
      let plainText = body.textContent.trim(); // Always have a plain text version
      if (platform === 'slack') {
        // Convert contenteditable HTML → Slack mrkdwn
        msg = convertNmBodyToSlackMrkdwn(body);
      } else {
        // Gmail: msg is plain text fallback, html_body has the rich HTML
        msg = plainText;
      }

      sendBtn.disabled=true; sendBtn.innerHTML='<span>Sending...</span>';
      try {
        const wn=platform==='slack'?'send_new_slack_message':'send_new_gmail_email';
        const payload={action:wn,platform,to:toR.map(r=>({id:r.id,name:r.name,email:r.email,type:r.type})),message:msg,attachments:nmPendingAttachments.length>0?nmPendingAttachments:undefined,timestamp:new Date().toISOString(),source,user_id:state.userData?.userId,authenticated:true};
        if (platform==='gmail') { payload.cc=ccR.map(r=>({id:r.id,name:r.name,email:r.email,type:r.type})); payload.subject=subIn.value.trim()||'(No Subject)'; payload.html_body=convertNmBodyToGmailHtml(body); }
        const res=await fetch(WEBHOOK_URL,{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify(payload)});
        if (res.ok) { if(typeof showToastNotification==='function') showToastNotification(`Message sent via ${platform==='slack'?'Slack':'Gmail'}!`); closeSlider(); }
        else throw new Error('Failed to send message');
      } catch(e) { console.error('Send error:',e); alert('Failed to send: '+e.message); sendBtn.disabled=false; sendBtn.innerHTML='<span>Send Message</span><span>➤</span>'; }
    });

    // --- Rich text: Bold/Italic shortcuts + Enter newline ---
    body.addEventListener('keydown', (e) => {
      // Bold: Cmd/Ctrl+B
      if ((e.metaKey || e.ctrlKey) && !e.shiftKey && (e.key === 'b' || e.key === 'B')) {
        e.preventDefault(); document.execCommand('bold', false, null); return;
      }
      // Italic: Cmd/Ctrl+I
      if ((e.metaKey || e.ctrlKey) && !e.shiftKey && (e.key === 'i' || e.key === 'I')) {
        e.preventDefault(); document.execCommand('italic', false, null); return;
      }
      // Cmd/Ctrl+K: Insert/edit hyperlink
      if ((e.metaKey || e.ctrlKey) && !e.shiftKey && (e.key === 'k' || e.key === 'K')) {
        e.preventDefault();
        _nmInsertLinkPrompt(body);
        return;
      }
    });

    // --- Paste URL on selected text → auto-hyperlink (all contenteditable areas) ---
    setupPasteToHyperlink(body);
    setupEmoticonReplace(body);

    setTimeout(()=>toIn.focus(),100);
  }

  // ============================================
  // Paste-to-Hyperlink: if text is selected and a URL is pasted, wrap selection as <a>
  // Exported so it can be called on any contenteditable input
  // ============================================
  const URL_REGEX = /^https?:\/\/[^\s]+$/i;

  function setupPasteToHyperlink(editableEl) {
    editableEl.addEventListener('paste', (e) => {
      const sel = window.getSelection();
      if (!sel.rangeCount) return;
      const range = sel.getRangeAt(0);

      // Only intercept if there's a selection (text highlighted)
      if (range.collapsed) return;

      const clipText = (e.clipboardData || window.clipboardData).getData('text/plain').trim();
      if (!clipText || !URL_REGEX.test(clipText)) return;

      // We have selected text + a URL in clipboard → create hyperlink
      e.preventDefault();
      e.stopImmediatePropagation();

      // Extract selected text, delete it, then insert an <a> wrapping it
      const selectedText = range.toString();
      range.deleteContents();
      const a = document.createElement('a');
      a.href = clipText;
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
      a.textContent = selectedText;
      range.insertNode(a);

      // Place cursor after the link
      const newRange = document.createRange();
      newRange.setStartAfter(a);
      newRange.collapse(true);
      sel.removeAllRanges();
      sel.addRange(newRange);
    });
  }

  // ============================================
  // Cmd+K: prompt user for URL and wrap selection
  // ============================================
  function _nmInsertLinkPrompt(editableEl) {
    const sel = window.getSelection();
    if (!sel.rangeCount) return;
    const range = sel.getRangeAt(0);
    const selectedText = range.toString();

    if (!selectedText) {
      // No selection — prompt for both text and URL
      const url = prompt('Enter URL:');
      if (!url) return;
      const text = prompt('Link text:', url) || url;
      const a = document.createElement('a');
      a.href = url; a.target = '_blank'; a.rel = 'noopener noreferrer';
      a.textContent = text;
      range.insertNode(a);
      // Add space after and place cursor
      const space = document.createTextNode('\u00A0');
      a.parentNode.insertBefore(space, a.nextSibling);
      const nr = document.createRange(); nr.setStartAfter(space); nr.collapse(true);
      sel.removeAllRanges(); sel.addRange(nr);
    } else {
      // Selection exists — just prompt for URL
      const url = prompt('Enter URL for "' + selectedText + '":');
      if (!url) return;
      document.execCommand('createLink', false, url);
      const newLink = sel.anchorNode?.parentElement?.closest('a') || editableEl.querySelector(`a[href="${url}"]`);
      if (newLink) { newLink.target = '_blank'; newLink.rel = 'noopener noreferrer'; }
      sel.collapseToEnd();
    }
  }

  // ============================================
  // Convert contenteditable HTML → Slack mrkdwn (with links, mentions, bold, italic)
  // ============================================
  function convertNmBodyToSlackMrkdwn(el) {
    const processNode = (node) => {
      if (node.nodeType === Node.TEXT_NODE) return node.textContent;
      if (node.nodeName === 'BR') return '\n';

      // Mention tags
      if (node.classList?.contains('mention-tag')) {
        const slackId = node.dataset.slackId;
        return slackId ? `<@${slackId}>` : node.textContent;
      }

      // Links → Slack format: <url|text>
      if (node.nodeName === 'A') {
        const href = node.getAttribute('href') || '';
        let inner = '';
        node.childNodes.forEach(c => { inner += processNode(c); });
        // If link text is same as URL, just use <url>
        if (inner.trim() === href.trim()) return `<${href}>`;
        return `<${href}|${inner}>`;
      }

      // Process children
      let inner = '';
      node.childNodes.forEach(c => { inner += processNode(c); });

      const tag = node.nodeName.toUpperCase();
      if (tag === 'B' || tag === 'STRONG') return `*${inner}*`;
      if (tag === 'I' || tag === 'EM') return `_${inner}_`;
      if (tag === 'S' || tag === 'STRIKE' || tag === 'DEL') return `~${inner}~`;
      if (tag === 'LI') {
        const parentList = node.closest('ol, ul');
        if (parentList?.nodeName === 'OL') {
          const items = Array.from(parentList.querySelectorAll(':scope > li'));
          return `${items.indexOf(node) + 1}. ${inner.trim()}`;
        }
        return `• ${inner.trim()}`;
      }
      if (tag === 'OL' || tag === 'UL') {
        const items = [];
        node.querySelectorAll(':scope > li').forEach(li => items.push(processNode(li)));
        return '\n' + items.join('\n') + '\n';
      }
      if (tag === 'DIV' || tag === 'P') return '\n' + inner;
      return inner;
    };
    let result = '';
    el.childNodes.forEach(n => { result += processNode(n); });
    return result.replace(/\n{3,}/g, '\n\n').trim();
  }

  // ============================================
  // Convert contenteditable HTML → Gmail-safe HTML
  // ============================================
  function convertNmBodyToGmailHtml(el) {
    // Clone and clean up the HTML for email
    const clone = el.cloneNode(true);
    // Convert mention tags to plain text for Gmail
    clone.querySelectorAll('.mention-tag').forEach(m => {
      const text = document.createTextNode(m.textContent);
      m.replaceWith(text);
    });
    // Ensure links have proper attributes
    clone.querySelectorAll('a').forEach(a => {
      a.target = '_blank';
      a.rel = 'noopener noreferrer';
    });
    // Convert div line breaks to <br> for email clients
    let html = clone.innerHTML;
    // Replace empty divs with <br>
    html = html.replace(/<div><br\s*\/?><\/div>/gi, '<br>');
    // Replace <div> wraps with <br> prefix (standard email line break)
    html = html.replace(/<div>/gi, '<br>').replace(/<\/div>/gi, '');
    // Clean up leading <br>
    html = html.replace(/^(<br\s*\/?>)+/, '');
    return html.trim();
  }

  // ============================================
  // Text Emoticon → Emoji auto-replace for contenteditable inputs
  // Triggers after space or Enter following a recognized emoticon
  // ============================================
  const EMOTICON_MAP = {
    ':)': '🙂', ':-)': '🙂', '(:': '🙂',
    ':D': '😀', ':-D': '😀',
    'xD': '😆', 'XD': '😆',
    ';)': '😉', ';-)': '😉',
    ':P': '😛', ':-P': '😛', ':p': '😛', ':-p': '😛',
    ':(': '😞', ':-(': '😞',
    ":'(": '😢', ":'-(": '😢',
    ':O': '😮', ':-O': '😮', ':o': '😮', ':-o': '😮',
    'B)': '😎', 'B-)': '😎',
    '<3': '❤️',
    '</3': '💔',
    ':*': '😘', ':-*': '😘',
    'O:)': '😇', 'O:-)': '😇',
    '>:(': '😠', '>:-(': '😠',
    ':/': '😕', ':-/': '😕', ':\\': '😕', ':-\\': '😕',
    ':S': '😖', ':-S': '😖',
    ':|': '😐', ':-|': '😐',
    '^_^': '😊',
    '-_-': '😑',
    'o_o': '😳', 'O_O': '😳',
    '>_<': '😣',
    ':thumbsup:': '👍', ':thumbsdown:': '👎',
    ':fire:': '🔥', ':heart:': '❤️', ':star:': '⭐',
    ':100:': '💯', ':pray:': '🙏', ':clap:': '👏',
    ':tada:': '🎉', ':rocket:': '🚀', ':eyes:': '👀',
    ':+1:': '👍', ':-1:': '👎', ':ok:': '👌',
    ':wave:': '👋', ':muscle:': '💪', ':brain:': '🧠',
    ':thinking:': '🤔', ':joy:': '😂', ':sob:': '😭',
    ':lol:': '😂', ':rofl:': '🤣',
  };

  // Build sorted keys (longest first so ":-)" matches before ":)")
  const EMOTICON_KEYS = Object.keys(EMOTICON_MAP).sort((a, b) => b.length - a.length);

  function setupEmoticonReplace(editableEl) {
    editableEl.addEventListener('input', (e) => {
      // Only trigger on space or newline insertion
      if (e.inputType !== 'insertText' || (e.data !== ' ' && e.data !== '\n')) return;

      const sel = window.getSelection();
      if (!sel.rangeCount) return;
      const range = sel.getRangeAt(0);
      const node = range.startContainer;
      if (node.nodeType !== Node.TEXT_NODE) return;

      const text = node.textContent;
      const cursorPos = range.startOffset;
      // Text before the space/newline that was just typed
      const beforeSpace = text.substring(0, cursorPos - 1);

      for (const emoticon of EMOTICON_KEYS) {
        if (beforeSpace.endsWith(emoticon)) {
          // Verify it's a word boundary (start of text or preceded by space/newline)
          const charBefore = beforeSpace.length > emoticon.length ? beforeSpace[beforeSpace.length - emoticon.length - 1] : null;
          if (charBefore !== null && charBefore !== ' ' && charBefore !== '\n' && charBefore !== '\u00A0') continue;

          const emoji = EMOTICON_MAP[emoticon];
          const emoticonStart = cursorPos - 1 - emoticon.length;
          // Replace emoticon + space with emoji + space
          node.textContent = text.substring(0, emoticonStart) + emoji + text.substring(cursorPos - 1);
          // Place cursor after emoji + space
          const newPos = emoticonStart + emoji.length + 1;
          const newRange = document.createRange();
          newRange.setStart(node, Math.min(newPos, node.textContent.length));
          newRange.collapse(true);
          sel.removeAllRanges();
          sel.addRange(newRange);
          break;
        }
      }
    });
  }

  window.OracleNewMessage = { showNewMessageSlider, setupPasteToHyperlink, setupEmoticonReplace };
})();
