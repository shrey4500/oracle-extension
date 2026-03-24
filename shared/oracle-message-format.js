// oracle-message-format.js — Message formatting, sanitization, email iframe rendering, attachments
// Used by transcript slider and FYI display in both newtab and sidepanel

(function () {
  'use strict';

  const { escapeHtml, createAuthenticatedPayload } = window.Oracle;

  const GMAIL_ATTACHMENT_WEBHOOK = 'https://n8n-kqq5.onrender.com/webhook/gmail-attachment';

  // ============================================
  // USER CHIP COLOR — consistent color from user_id
  // ============================================
  const chipPalette = [
    '#4a5bc7', '#d9434e', '#2d7fc4', '#2d9d5e', '#c4487a',
    '#7b5ea7', '#9b4ec4', '#3a7bbf', '#c46a3a', '#3a6ec4',
    '#9b3aaf', '#c43a8f'
  ];
  // Brighter variants for dark mode legibility
  const chipPaletteDark = [
    '#7b8ef8', '#f47a84', '#5eb3f7', '#5dd98e', '#f47aa6',
    '#b48fdd', '#c97ef7', '#6baef2', '#f49a6a', '#6b9ef7',
    '#c96be0', '#f76abf'
  ];
  function chipColorFromId(uid) {
    let hash = 0;
    const str = String(uid || 'U');
    for (let i = 0; i < str.length; i++) {
      hash = ((hash << 5) - hash) + str.charCodeAt(i);
      hash = hash & hash;
    }
    return Math.abs(hash) % chipPalette.length;
  }
  function buildMentionChip(uid, name) {
    const idx = chipColorFromId(uid);
    const isDark = document.body?.classList?.contains('dark-mode') || false;
    const color = isDark ? chipPaletteDark[idx] : chipPalette[idx];
    // Convert hex to rgb for proper rgba background
    const r = parseInt(color.slice(1,3), 16);
    const g = parseInt(color.slice(3,5), 16);
    const b = parseInt(color.slice(5,7), 16);
    return `<span data-mention-chip="1" style="--chip-color:${color};display:inline-flex;align-items:center;background:rgba(${r},${g},${b},${isDark ? '0.2' : '0.15'});color:${color};padding:2px 8px;border-radius:10px;font-size:12px;font-weight:600;white-space:nowrap;vertical-align:baseline;line-height:1.4;">@${name}</span>`;
  }

  // ============================================
  // EMOJI MAP (Slack codes → Unicode)
  // ============================================
  const emojiMap = {
    // Faces — smiling
    ':slightly_smiling_face:':'🙂',':smile:':'😄',':grinning:':'😀',':laughing:':'😆',':blush:':'😊',
    ':smiley:':'😃',':grin:':'😁',':rofl:':'🤣',':relaxed:':'☺️',':yum:':'😋',
    ':smiling_face_with_3_hearts:':'🥰',':smiling_face_with_tear:':'🥲',':upside_down_face:':'🙃',
    ':melting_face:':'🫠',':star_struck:':'🤩',':partying_face:':'🥳',
    // Faces — affection
    ':wink:':'😉',':heart_eyes:':'😍',':kissing_heart:':'😘',':kissing:':'😗',
    ':kissing_smiling_eyes:':'😙',':kissing_closed_eyes:':'😚',
    // Faces — tongue
    ':stuck_out_tongue:':'😛',':stuck_out_tongue_winking_eye:':'😜',':stuck_out_tongue_closed_eyes:':'😝',
    ':zany_face:':'🤪',':face_with_tongue:':'😛',':money_mouth_face:':'🤑',
    // Faces — hands
    ':hugging_face:':'🤗',':shushing_face:':'🤫',':thinking_face:':'🤔',':thinking:':'🤔',
    ':face_with_hand_over_mouth:':'🤭',':saluting_face:':'🫡',':zipper_mouth_face:':'🤐',
    // Faces — neutral / skeptical
    ':raised_eyebrow:':'🤨',':neutral_face:':'😐',':expressionless:':'😑',':no_mouth:':'😶',
    ':dotted_line_face:':'🫥',':face_in_clouds:':'😶‍🌫️',':smirk:':'😏',':unamused:':'😒',
    ':rolling_eyes:':'🙄',':grimacing:':'😬',':face_exhaling:':'😮‍💨',':lying_face:':'🤥',
    // Faces — sleepy
    ':relieved:':'😌',':pensive:':'😔',':sleepy:':'😪',':drooling_face:':'🤤',':sleeping:':'😴',
    // Faces — unwell
    ':mask:':'😷',':face_with_thermometer:':'🤒',':face_with_head_bandage:':'🤕',':nauseated_face:':'🤢',
    ':face_vomiting:':'🤮',':sneezing_face:':'🤧',':hot_face:':'🥵',':cold_face:':'🥶',
    ':woozy_face:':'🥴',':dizzy_face:':'😵',':face_with_spiral_eyes:':'😵‍💫',':exploding_head:':'🤯',
    // Faces — concerned
    ':confused:':'😕',':worried:':'😟',':slightly_frowning_face:':'🙁',':frowning_face:':'☹️',
    ':open_mouth:':'😮',':hushed:':'😯',':astonished:':'😲',':flushed:':'😳',
    ':pleading_face:':'🥺',':face_holding_back_tears:':'🥹',':anguished:':'😧',':fearful:':'😨',
    ':cold_sweat:':'😰',':disappointed_relieved:':'😥',':sweat:':'😓',
    ':persevere:':'😣',':disappointed:':'😞',':confounded:':'😖',':tired_face:':'😫',':weary:':'😩',
    // Faces — negative
    ':triumph:':'😤',':angry:':'😠',':rage:':'😡',':cursing_face:':'🤬',':face_with_symbols_on_mouth:':'🤬',
    // Faces — crying
    ':cry:':'😢',':sob:':'😭',':sweat_smile:':'😅',':joy:':'😂',
    // Faces — costume
    ':poop:':'💩',':ghost:':'👻',':skull:':'💀',':skull_and_crossbones:':'☠️',':alien:':'👽',
    ':clown_face:':'🤡',':imp:':'👿',':smiling_imp:':'😈',':japanese_ogre:':'👹',':japanese_goblin:':'👺',
    // Faces — glasses & hats
    ':nerd_face:':'🤓',':sunglasses:':'😎',':disguised_face:':'🥸',':cowboy_hat_face:':'🤠',
    // Faces — other
    ':innocent:':'😇',':scream:':'😱',':face_with_monocle:':'🧐',':shrug:':'🤷',
    ':man_shrugging:':'🤷‍♂️',':woman_shrugging:':'🤷‍♀️',':face_palm:':'🤦',':facepalm:':'🤦',
    // Hands
    ':raised_hands:':'🙌',':clap:':'👏',':pray:':'🙏',':thumbsup:':'👍',':thumbsdown:':'👎',
    ':+1:':'👍',':-1:':'👎',':ok_hand:':'👌',':wave:':'👋',':point_up:':'☝️',
    ':point_down:':'👇',':point_left:':'👈',':point_right:':'👉',
    ':handshake:':'🤝',':fist:':'✊',':v:':'✌️',':crossed_fingers:':'🤞',':pinched_fingers:':'🤌',
    ':call_me_hand:':'🤙',':metal:':'🤘',':pinching_hand:':'🤏',':writing_hand:':'✍️',
    ':palms_up_together:':'🤲',':open_hands:':'👐',':muscle:':'💪',
    // Hearts & symbols
    ':heart:':'❤️',':broken_heart:':'💔',':orange_heart:':'🧡',':yellow_heart:':'💛',
    ':green_heart:':'💚',':blue_heart:':'💙',':purple_heart:':'💜',':black_heart:':'🖤',
    ':white_heart:':'🤍',':brown_heart:':'🤎',':sparkling_heart:':'💖',':two_hearts:':'💕',
    ':revolving_hearts:':'💞',':heartbeat:':'💓',':heartpulse:':'💗',':growing_heart:':'💗',
    ':cupid:':'💘',':heart_on_fire:':'❤️‍🔥',':mending_heart:':'❤️‍🩹',
    ':fire:':'🔥',':star:':'⭐',':star2:':'🌟',':sparkles:':'✨',':100:':'💯',
    ':warning:':'⚠️',':zap:':'⚡',':boom:':'💥',':collision:':'💥',':dizzy:':'💫',
    // Status & symbols
    ':white_check_mark:':'✅',':x:':'❌',':heavy_check_mark:':'✔️',':question:':'❓',':exclamation:':'❗',
    ':red_circle:':'🔴',':large_blue_circle:':'🔵',':green_circle:':'🟢',':yellow_circle:':'🟡',
    ':white_circle:':'⚪',':black_circle:':'⚫',':arrow_right:':'➡️',':arrow_left:':'⬅️',
    ':arrow_up:':'⬆️',':arrow_down:':'⬇️',':heavy_plus_sign:':'➕',':heavy_minus_sign:':'➖',
    // Animals & nature
    ':eyes:':'👀',':see_no_evil:':'🙈',':hear_no_evil:':'🙉',':speak_no_evil:':'🙊',
    ':dog:':'🐶',':cat:':'🐱',':bear:':'🐻',':unicorn:':'🦄',':butterfly:':'🦋',
    ':bee:':'🐝',':bug:':'🐛',':snake:':'🐍',':turtle:':'🐢',':octopus:':'🐙',
    ':sunflower:':'🌻',':rose:':'🌹',':cherry_blossom:':'🌸',':four_leaf_clover:':'🍀',
    ':evergreen_tree:':'🌲',':deciduous_tree:':'🌳',':cactus:':'🌵',':palm_tree:':'🌴',
    // Objects
    ':rocket:':'🚀',':tada:':'🎉',':party_popper:':'🎉',':balloon:':'🎈',':confetti_ball:':'🎊',
    ':gift:':'🎁',':ribbon:':'🎀',':bulb:':'💡',':memo:':'📝',':pencil:':'✏️',
    ':pushpin:':'📌',':calendar:':'📅',':clock:':'🕐',':hourglass:':'⏳',':email:':'📧',
    ':envelope:':'✉️',':phone:':'📞',':computer:':'💻',':link:':'🔗',':lock:':'🔒',
    ':key:':'🔑',':hammer:':'🔨',':wrench:':'🔧',':gear:':'⚙️',':shield:':'🛡️',
    ':chart_with_upwards_trend:':'📈',':chart_with_downwards_trend:':'📉',':bar_chart:':'📊',
    ':moneybag:':'💰',':dollar:':'💵',':credit_card:':'💳',':gem:':'💎',':trophy:':'🏆',
    ':medal:':'🏅',':crown:':'👑',':brain:':'🧠',':robot_face:':'🤖',':robot:':'🤖',
    ':zzz:':'💤',':speech_balloon:':'💬',':thought_balloon:':'💭',':loudspeaker:':'📢',
    ':bell:':'🔔',':no_bell:':'🔕',':microphone:':'🎤',':headphones:':'🎧',':musical_note:':'🎵',
    ':camera:':'📷',':video_camera:':'📹',':tv:':'📺',':radio:':'📻',':satellite:':'📡',
    ':book:':'📖',':books:':'📚',':newspaper:':'📰',':clipboard:':'📋',':file_folder:':'📁',
    ':inbox_tray:':'📥',':outbox_tray:':'📤',':package:':'📦',':mailbox:':'📬',
    // Food
    ':coffee:':'☕',':tea:':'🍵',':pizza:':'🍕',':hamburger:':'🍔',':fries:':'🍟',
    ':cookie:':'🍪',':cake:':'🎂',':ice_cream:':'🍨',':beer:':'🍺',':wine_glass:':'🍷',
    ':tropical_drink:':'🍹',':champagne:':'🍾',':apple:':'🍎',':banana:':'🍌',':watermelon:':'🍉',
    // Travel & weather
    ':sunny:':'☀️',':cloud:':'☁️',':umbrella:':'☂️',':snowflake:':'❄️',':rainbow:':'🌈',
    ':earth_americas:':'🌎',':earth_africa:':'🌍',':earth_asia:':'🌏',':globe_with_meridians:':'🌐',
    ':airplane:':'✈️',':car:':'🚗',':bus:':'🚌',':bike:':'🚲',':ship:':'🚢',':house:':'🏠',
    // Flags & misc
    ':flag-us:':'🇺🇸',':flag-gb:':'🇬🇧',':flag-in:':'🇮🇳',':flag-jp:':'🇯🇵',':flag-de:':'🇩🇪',
    ':checkered_flag:':'🏁',':triangular_flag_on_post:':'🚩',':crossed_flags:':'🎌',':pirate_flag:':'🏴‍☠️'
  };

  // ============================================
  // truncateUrl — Show clean, shortened URL for display
  // Extracts domain + path, resolves Google redirects, caps at ~60 chars
  // ============================================
  function truncateUrl(url) {
    if (!url || url.length <= 60) return url;
    try {
      const u = new URL(url);
      const domain = u.hostname.replace(/^www\./, '');
      const path = u.pathname;
      // If path is short enough, show domain + path
      if ((domain + path).length <= 55) {
        return domain + path + (u.search ? '…' : '');
      }
      // Otherwise show domain + truncated path
      const pathParts = path.split('/').filter(Boolean);
      if (pathParts.length <= 1) {
        return domain + '/' + (pathParts[0] || '').substring(0, 30) + '…';
      }
      // Show domain + first path segment + ... + last segment (truncated)
      const first = pathParts[0];
      const last = pathParts[pathParts.length - 1];
      const truncated = domain + '/' + first + '/…/' + (last.length > 20 ? last.substring(0, 20) + '…' : last);
      return truncated.length > 65 ? domain + '/…/' + (last.length > 25 ? last.substring(0, 25) + '…' : last) : truncated;
    } catch {
      // Fallback: simple character truncation
      return url.substring(0, 55) + '…';
    }
  }

  // ============================================
  // formatMessageContent — Slack/plain text formatting
  // ============================================
  function formatMessageContent(text) {
    if (!text) return '';

    // Pre-process Slack link patterns BEFORE HTML detection
    // These look like HTML tags (<url|text>, <mailto:...>) but are Slack formatting
    let processed = text;

    // New format: <<@UID|DisplayName>> from updated construct-transcript
    // Preserve user_id for consistent chip coloring
    processed = processed.replace(/<<@([A-Z0-9]+)\|([^>]+)>>/g, '%%UMENTION:$1%%$2%%ENDUMENTION%%');

    // Legacy Slack user mentions <@U1234|Name> or <@U1234> → styled chip placeholder
    processed = processed.replace(/<@([A-Z0-9]+)\|([^>]+)>/g, '%%UMENTION:$1%%$2%%ENDUMENTION%%');
    processed = processed.replace(/<@([A-Z0-9]+)>/g, '%%UMENTION:$1%%$1%%ENDUMENTION%%');
    // Slack channel mentions <!subteam^ID|@Name> or <!here> etc
    processed = processed.replace(/<!subteam\^[A-Z0-9]+\|@([^>]+)>/g, '%%MENTION%%$1%%ENDMENTION%%');
    processed = processed.replace(/<!([a-z]+)>/g, '%%MENTION%%$1%%ENDMENTION%%');

    // Slack <url|text> → markdown-style placeholder to preserve through HTML detection
    processed = processed.replace(/<(https?:\/\/[^|>]+)\|([^>]+)>/g, '%%SLACKLINK:$1%%$2%%ENDSLACKLINK%%');
    // Slack <mailto:email|text>
    processed = processed.replace(/<mailto:([^|>]+)\|([^>]+)>/g, '%%SLACKMAILTO:$1%%$2%%ENDSLACKMAILTO%%');
    // Slack <mailto:email>
    processed = processed.replace(/<mailto:([^>]+)>/g, '%%SLACKMAILTO:$1%%$1%%ENDSLACKMAILTO%%');
    // Slack <url> (bare)
    processed = processed.replace(/<(https?:\/\/[^>]+)>/g, '%%SLACKLINK:$1%%$1%%ENDSLACKLINK%%');

    // Detect actual HTML content — require known HTML tags or self-closing patterns
    // Avoid false positives from angle brackets in plain text like <Product Account Id>
    const hasRealHtml = /<(?:div|span|p|br|a |img |table|tr|td|th|ul|ol|li|h[1-6]|b|i|u|em|strong|pre|code|blockquote|hr|head|body|html|style|font|section|article|header|footer|main|aside|nav|form|input|button|label|select|textarea|meta|link)[\s>\/]/i.test(processed);
    if (hasRealHtml) {
      // Restore Slack links before sanitizing
      processed = processed.replace(/%%SLACKLINK:(.*?)%%(.*?)%%ENDSLACKLINK%%/g, (_, url, label) => {
        const displayText = /^https?:\/\//.test(label) ? truncateUrl(label) : label;
        return `<a href="${url}" target="_blank" style="color:#667eea;text-decoration:underline;word-break:break-all;">${displayText}</a>`;
      });
      processed = processed.replace(/%%SLACKMAILTO:(.*?)%%(.*?)%%ENDSLACKMAILTO%%/g, '<a href="mailto:$1" style="color:#667eea;text-decoration:underline;">$2</a>');
      // Also convert Markdown links in HTML content
      processed = processed.replace(/\[([^\]]+)\]\((https?:\/\/[^)]+)\)/g, (_, label, url) => {
        const displayText = /^https?:\/\//.test(label) ? truncateUrl(label) : label;
        return `<a href="${url}" target="_blank" style="color:#667eea;text-decoration:underline;word-break:break-all;">${displayText}</a>`;
      });
      processed = processed.replace(/\[([^\]]+)\]\(mailto:([^)]+)\)/g, '<a href="mailto:$2" style="color:#667eea;text-decoration:underline;">$1</a>');
      // Restore @mention chips in HTML content — with user_id colors
      processed = processed.replace(/%%UMENTION:([A-Z0-9]+)%%(.*?)%%ENDUMENTION%%/g, (_, uid, name) => buildMentionChip(uid, name));
      processed = processed.replace(/%%MENTION%%(.*?)%%ENDMENTION%%/g, '<span style="display:inline-flex;align-items:center;gap:2px;background:linear-gradient(135deg,rgba(102,126,234,0.15),rgba(118,75,162,0.1));color:#667eea;padding:1px 8px;border-radius:10px;font-size:12px;font-weight:600;white-space:nowrap;vertical-align:baseline;">@$1</span>');
      return sanitizeHtml(processed);
    }

    // Restore Slack links for plain text path (they'll be handled below)
    processed = processed.replace(/%%SLACKLINK:(.*?)%%(.*?)%%ENDSLACKLINK%%/g, '<$1|$2>');
    processed = processed.replace(/%%SLACKMAILTO:(.*?)%%(.*?)%%ENDSLACKMAILTO%%/g, '<mailto:$1|$2>');

    // Plain text
    let fmt = processed;

    // Decode HTML entities
    fmt = fmt.replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&amp;/g,'&').replace(/&quot;/g,'"').replace(/&#39;/g,"'");

    // Convert Slack link patterns to placeholders BEFORE escaping HTML
    // Slack rich_text block format: [display_text]\n(actual_url) — display text in brackets, URL in parens on next line
    fmt = fmt.replace(/\[([^\]]*(?:\[[^\]]*\][^\]]*)*)\]\s*\n\s*\((https?:\/\/[^)]+)\)/g, (_, label, url) => {
      const displayText = /^https?:\/\//.test(label) ? truncateUrl(url) : label;
      return `%%LINK_A%%${url}%%LINK_B%%${escapeHtml(displayText)}%%LINK_T%%${escapeHtml(url)}%%LINK_END%%`;
    });
    // Slack rich_text: standalone [url] in square brackets (no following paren URL)
    fmt = fmt.replace(/\[(https?:\/\/[^\]\s]+(?:\[[^\]]*\][^\]\s]*)*)\](?!\s*\n?\s*\()/g, (_, url) => {
      return `%%LINK_A%%${url}%%LINK_B%%${truncateUrl(url)}%%LINK_T%%${escapeHtml(url)}%%LINK_END%%`;
    });
    // Slack rich_text: standalone (url) in parentheses on its own
    fmt = fmt.replace(/(?:^|\n)\s*\((https?:\/\/[^)]+)\)/g, (match, url) => {
      const prefix = match.startsWith('\n') ? '\n' : '';
      return `${prefix}%%LINK_A%%${url}%%LINK_B%%${truncateUrl(url)}%%LINK_T%%${escapeHtml(url)}%%LINK_END%%`;
    });
    // Slack <url|text>
    fmt = fmt.replace(/<(https?:\/\/[^|>]+)\|([^>]+)>/g, (_, url, label) => {
      const displayText = /^https?:\/\//.test(label) ? truncateUrl(label) : label;
      return `%%LINK_A%%${url}%%LINK_B%%${escapeHtml(displayText)}%%LINK_T%%${escapeHtml(label)}%%LINK_END%%`;
    });
    // Slack <mailto:email|text>
    fmt = fmt.replace(/<mailto:([^|>]+)\|([^>]+)>/g, '%%MAILTO_A%%$1%%MAILTO_B%%$2%%MAILTO_END%%');
    // Slack <mailto:email>
    fmt = fmt.replace(/<mailto:([^>]+)>/g, '%%MAILTO_A%%$1%%MAILTO_B%%$1%%MAILTO_END%%');
    // Slack <url> bare
    fmt = fmt.replace(/<(https?:\/\/[^>]+)>/g, (_, url) => {
      return `%%LINK_A%%${url}%%LINK_B%%${truncateUrl(url)}%%LINK_T%%${escapeHtml(url)}%%LINK_END%%`;
    });

    // Markdown links [text](url) — handle BEFORE escapeHtml to avoid & encoding issues
    fmt = fmt.replace(/\[([^\]]+)\]\((https?:\/\/[^)]+)\)/g, (_, label, url) => {
      const displayText = /^https?:\/\//.test(label) ? truncateUrl(label) : escapeHtml(label);
      return `%%LINK_A%%${url}%%LINK_B%%${displayText}%%LINK_T%%${escapeHtml(label)}%%LINK_END%%`;
    });
    // Markdown mailto links [text](mailto:email)
    fmt = fmt.replace(/\[([^\]]+)\]\(mailto:([^)]+)\)/g, '%%MAILTO_A%%$2%%MAILTO_B%%$1%%MAILTO_END%%');

    // Plain URLs — convert BEFORE escapeHtml to avoid & encoding issues
    fmt = fmt.replace(/(^|[\s\n])(https?:\/\/[^\s<>)\]]+)/g, (_, prefix, url) => {
      return `${prefix}%%LINK_A%%${url}%%LINK_B%%${truncateUrl(url)}%%LINK_T%%${escapeHtml(url)}%%LINK_END%%`;
    });

    // Escape HTML (safe now that all links are placeholdered)
    fmt = escapeHtml(fmt);

    // Restore Slack link placeholders as proper <a> tags
    fmt = fmt.replace(/%%LINK_A%%(.*?)%%LINK_B%%(.*?)%%LINK_T%%(.*?)%%LINK_END%%/g,
      '<a href="$1" target="_blank" style="color:#667eea;text-decoration:underline;word-break:break-all;" title="$3">$2</a>');
    fmt = fmt.replace(/%%MAILTO_A%%(.*?)%%MAILTO_B%%(.*?)%%MAILTO_END%%/g,
      '<a href="mailto:$1" style="color:#667eea;text-decoration:underline;">$2</a>');

    // Numbered headers (run BEFORE mention restoration to avoid matching colons in HTML attributes)
    // Only if the colon isn't part of a URL like https: or inside a %%PLACEHOLDER%%
    fmt = fmt.replace(/^(\d+\.\s*[^:%\n]+:)(?!\/\/)/gm, '<strong style="font-weight:600;color:var(--text-primary, #2c3e50);display:block;margin-top:12px;margin-bottom:4px;">$1</strong>');
    // Numbered list items
    fmt = fmt.replace(/^(\d+)\.\s+/gm, '<span style="color:#667eea;font-weight:600;margin-right:4px;">$1.</span> ');
    // Bold *text*
    fmt = fmt.replace(/(?<![\\w*])\*(\S[^*\n]{0,148}?\S)\s?\*(?![\\w*])/g, '<strong style="font-weight:700;">$1</strong>');
    // Italic _text_
    fmt = fmt.replace(/(?<!\w)_([^_\n]+)_(?!\w)/g, '<em style="font-style:italic;">$1</em>');
    // Strikethrough ~text~
    fmt = fmt.replace(/~([^~\n]+)~/g, '<del style="text-decoration:line-through;opacity:0.7;">$1</del>');

    // Restore @mention chips — with user_id-based colors
    fmt = fmt.replace(/%%UMENTION:([A-Z0-9]+)%%(.*?)%%ENDUMENTION%%/g, (_, uid, name) => buildMentionChip(uid, name));
    fmt = fmt.replace(/%%MENTION%%(.*?)%%ENDMENTION%%/g, '<span style="display:inline-flex;align-items:center;gap:2px;background:linear-gradient(135deg,rgba(102,126,234,0.15),rgba(118,75,162,0.1));color:#667eea;padding:1px 8px;border-radius:10px;font-size:12px;font-weight:600;white-space:nowrap;vertical-align:baseline;">@$1</span>');
    // Plain @Name mentions are left as-is; only <<@UID|Name>> from n8n are rendered as chips

    // Slack emoji codes
    Object.keys(emojiMap).forEach(code => { fmt = fmt.split(code).join(emojiMap[code]); });
    // Replace remaining emoji shortcodes with map lookup, fallback heuristics, or styled placeholder
    fmt = fmt.replace(/:([a-z0-9_+-]+):/g, (match, name) => {
      const full = `:${name}:`;
      if (emojiMap[full]) return emojiMap[full];

      // Strip skin-tone suffix: e.g. "thumbsup::skin-tone-3" or "pray::skin-tone-2"
      const skinToneMatch = name.match(/^(.+?)(?:::?skin-tone-\d)$/);
      if (skinToneMatch) {
        const base = `:${skinToneMatch[1]}:`;
        if (emojiMap[base]) return emojiMap[base];
      }

      // Strip trailing numbers: e.g. "welcome2" → "welcome", "tada2" → "tada"
      const numStripped = name.replace(/\d+$/, '');
      if (numStripped && numStripped !== name) {
        const base = `:${numStripped}:`;
        if (emojiMap[base]) return emojiMap[base];
      }

      // Common Slack custom emoji fallbacks (workspace-specific emojis that have standard equivalents)
      const customFallbacks = {
        'welcome':'👋', 'thanks':'🙏', 'thankyou':'🙏', 'thank_you':'🙏', 'thank-you':'🙏',
        'congrats':'🎉', 'congratulations':'🎉', 'celebrate':'🎉', 'yay':'🎉',
        'love':'❤️', 'lgtm':'👍', 'shipit':'🚀', 'ship':'🚀', 'deploy':'🚀',
        'approved':'✅', 'done':'✅', 'complete':'✅', 'yes':'✅', 'no':'❌',
        'ack':'👍', 'noted':'📝', 'alert':'🚨', 'urgent':'🚨', 'help':'🆘',
        'coffee':'☕', 'beer':'🍺', 'lunch':'🍽️', 'food':'🍕',
        'plus':'➕', 'minus':'➖', 'check':'✅', 'info':'ℹ️',
        'wave':'👋', 'hi':'👋', 'hello':'👋', 'bye':'👋',
        'up':'⬆️', 'down':'⬇️', 'left':'⬅️', 'right':'➡️',
        'fast':'⚡', 'slow':'🐢', 'bug':'🐛', 'fix':'🔧',
        'idea':'💡', 'tip':'💡', 'question':'❓', 'answer':'💬',
        'great':'🔥', 'awesome':'🔥', 'nice':'👍', 'good':'👍', 'cool':'😎',
        'sad':'😢', 'happy':'😊', 'wow':'😮', 'surprised':'😮',
        'party':'🎉', 'dance':'💃', 'music':'🎵', 'sing':'🎤',
        'blob_wave':'👋', 'blob_clap':'👏', 'blob_thumbsup':'👍', 'blob_dance':'💃',
        'blob_heart':'❤️', 'blob_thinking':'🤔', 'blob_eyes':'👀',
        'parrot':'🦜', 'meow':'🐱', 'doge':'🐕',
        // Animated / workspace custom emoji fallbacks
        'gifire':'🔥', 'fire-but-animated':'🔥', 'fire_gif':'🔥', 'fireflame':'🔥',
        'parrotdoge':'🦜', 'partydoge':'🎉', 'partyparrot':'🦜', 'party_parrot':'🦜',
        'fastparrot':'🦜', 'slowparrot':'🦜', 'shuffleparrot':'🦜', 'congaparrot':'🦜',
        'blobdance':'💃', 'blob_dance':'💃', 'blobjam':'🎵', 'catjam':'🐱',
        'meow_party':'🐱', 'nyan':'🐱', 'nyancat':'🐱',
        'dancingbanana':'🍌', 'bananadance':'🍌',
        'loadingdots':'⏳', 'loading':'⏳', 'spinner':'⏳',
        'stonks':'📈', 'notstonks':'📉',
        'this':'👆', 'that':'👆', 'upvote':'👍', 'downvote':'👎',
        'applause':'👏', 'clapping':'👏', 'slow_clap':'👏', 'clap':'👏',
        'bow':'🙇', 'salute':'🫡', 'facepalm':'🤦', 'shrug':'🤷',
      };
      // Try custom fallback directly
      if (customFallbacks[name]) return customFallbacks[name];
      // Try with numbers stripped
      if (numStripped && customFallbacks[numStripped]) return customFallbacks[numStripped];

      // Final fallback: render as a subtle inline tag so it doesn't look broken
      return `<span style="display:inline-block;background:rgba(102,126,234,0.08);color:var(--text-muted,#7f8c8d);padding:1px 6px;border-radius:4px;font-size:11px;vertical-align:baseline;">:${name}:</span>`;
    });

    // Code blocks (multiline: ```...``` with dotAll flag)
    fmt = fmt.replace(/```([\s\S]+?)```/g,
      '<pre style="background:rgba(45,55,72,0.08);border:1px solid rgba(45,55,72,0.15);border-radius:6px;padding:12px;margin:8px 0;font-family:\'SF Mono\',Monaco,\'Courier New\',monospace;font-size:12px;overflow-x:hidden;white-space:pre-wrap;word-break:break-all;max-width:100%;">$1</pre>');
    // Inline code
    fmt = fmt.replace(/`([^`\n]+)`/g,
      '<code style="background:rgba(45,55,72,0.08);border-radius:4px;padding:2px 6px;font-family:\'SF Mono\',Monaco,\'Courier New\',monospace;font-size:12px;color:#e53e3e;">$1</code>');

    // Emails (still needs to run post-escape since it's simple pattern matching)
    fmt = fmt.replace(/(?<!href="mailto:|">)(?<![a-zA-Z0-9._%+-])([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})(?!<\/a>)/g, '<a href="mailto:$1" style="color:#667eea;text-decoration:underline;">$1</a>');

    // [Formatting regexes moved to run before mention restoration — see above]

    // Remove [image: ...] placeholders
    fmt = fmt.replace(/\[image:[^\]]+\]/g, '');

    // Convert newlines to <br> BEFORE blockquote processing (except inside <pre> blocks)
    const prePartsEarly = fmt.split(/(<pre[^>]*>[\s\S]*?<\/pre>)/);
    fmt = prePartsEarly.map(p => p.startsWith('<pre') ? p : p.replace(/\n/g, '<br>')).join('');

    // Blockquotes (> prefix) — split on <br> now that newlines are converted
    const lines = fmt.split('<br>');
    let inQuote = false, quoteLines = [], result = [];
    for (const line of lines) {
      if (/^(&gt;|>)\s?/.test(line.trim())) {
        quoteLines.push(line.trim().replace(/^(&gt;|>)\s?/, ''));
        inQuote = true;
      } else {
        if (inQuote && quoteLines.length > 0) {
          result.push(`<blockquote style="border-left:4px solid #667eea;margin:8px 0;padding:8px 12px;background:rgba(102,126,234,0.08);border-radius:0 8px 8px 0;color:inherit;font-style:normal;">${quoteLines.join('<br>')}</blockquote>`);
          quoteLines = []; inQuote = false;
        }
        result.push(line);
      }
    }
    if (quoteLines.length > 0) {
      result.push(`<blockquote style="border-left:4px solid #667eea;margin:8px 0;padding:8px 12px;background:rgba(102,126,234,0.08);border-radius:0 8px 8px 0;color:inherit;font-style:normal;">${quoteLines.join('<br>')}</blockquote>`);
    }
    fmt = result.join('<br>');

    // Cleanup
    fmt = fmt.replace(/<br>{3,}/g, '<br><br>');
    // Newlines outside <pre> blocks
    const preParts = fmt.split(/(<pre[^>]*>[\s\S]*?<\/pre>)/);
    fmt = preParts.map(p => p.startsWith('<pre') ? p : p.replace(/\n/g, '<br>')).join('');

    return fmt;
  }

  // ============================================
  // sanitizeHtml — Strip dangerous elements, style safe elements
  // ============================================
  function sanitizeHtml(html) {
    const temp = document.createElement('div');
    temp.innerHTML = html;

    temp.querySelectorAll('script, style, iframe, object, embed').forEach(el => el.remove());

    temp.querySelectorAll('*').forEach(el => {
      Array.from(el.attributes).forEach(attr => {
        if (attr.name.startsWith('on') || (attr.name === 'href' && attr.value.startsWith('javascript:'))) {
          el.removeAttribute(attr.name);
        }
      });
    });

    temp.querySelectorAll('table').forEach(t => { t.style.cssText = 'border-collapse:collapse;width:100%;margin:10px 0;font-size:13px;overflow-x:hidden;display:block;max-width:100%;table-layout:fixed;'; });
    temp.querySelectorAll('th').forEach(th => { th.style.cssText = 'background:linear-gradient(45deg,#667eea,#764ba2);color:white;padding:10px 12px;text-align:left;font-weight:600;border:1px solid rgba(102,126,234,0.3);white-space:nowrap;'; });
    temp.querySelectorAll('td').forEach(td => { td.style.cssText = 'padding:8px 12px;border:1px solid rgba(225,232,237,0.6);vertical-align:top;'; });
    temp.querySelectorAll('tr').forEach((tr, i) => { if (i % 2 === 1) tr.style.backgroundColor = 'rgba(102,126,234,0.03)'; });
    temp.querySelectorAll('a').forEach(a => { a.style.cssText = 'color:#667eea;text-decoration:underline;'; a.setAttribute('target', '_blank'); });
    temp.querySelectorAll('p').forEach(p => { p.style.cssText = 'margin:8px 0;line-height:1.5;'; });
    temp.querySelectorAll('b, strong').forEach(b => { b.style.cssText = 'font-weight:600;color:#2c3e50;'; });
    temp.querySelectorAll('li').forEach(li => { li.style.whiteSpace = 'normal'; li.style.overflowWrap = 'break-word'; li.style.wordBreak = 'break-word'; });
    // Strip white-space:pre from all elements (Gmail sets this on many elements causing overflow)
    temp.querySelectorAll('*').forEach(el => {
      const style = el.getAttribute('style') || '';
      if (style.includes('white-space') && (style.includes('pre') || style.includes('nowrap'))) {
        el.style.whiteSpace = 'normal';
      }
    });

    return temp.innerHTML;
  }

  // ============================================
  // isComplexEmailHtml — Detect complex email (nested tables, signatures)
  // ============================================
  function isComplexEmailHtml(html) {
    if (!html) return false;
    const temp = document.createElement('div');
    temp.innerHTML = html;
    const nestedTables = temp.querySelectorAll('table table');
    const sigs = temp.querySelectorAll('[id*="Signature"], [id*="signature"], .elementToProof');
    const hrs = temp.querySelectorAll('hr');
    const hasMany = (html.match(/style="/g) || []).length > 10;
    return nestedTables.length > 0 || sigs.length > 0 || (hrs.length > 0 && hasMany);
  }

  // ============================================
  // stripEmailQuotedContent — Remove quoted email chain content
  // ============================================
  function stripEmailQuotedContent(html) {
    if (!html) return html;
    const temp = document.createElement('div');
    temp.innerHTML = html;

    // Remove Gmail quoted content containers
    temp.querySelectorAll('.gmail_quote, .gmail_quote_container').forEach(el => el.remove());

    // Remove Outlook-style quoted content (reference messages)
    temp.querySelectorAll('#m_[id*="mail-editor-reference-message-container"], [id*="reference-message-container"]').forEach(el => el.remove());
    temp.querySelectorAll('div[id*="mail-editor-reference-message"]').forEach(el => el.remove());

    // Remove email signatures
    temp.querySelectorAll('.gmail_signature, .gmail_signature_prefix').forEach(el => el.remove());

    // Remove confidentiality disclaimers (common in enterprise emails)
    temp.querySelectorAll('b > i > span').forEach(el => {
      if (el.textContent && el.textContent.includes('confidential') && el.textContent.length > 100) {
        // Remove the entire bold/italic wrapper
        const parent = el.closest('b');
        if (parent) parent.remove();
      }
    });

    // Remove trailing <br> and empty divs
    const result = temp.innerHTML
      .replace(/(<br\s*\/?>[\s\n]*)+$/gi, '')
      .replace(/(<div>\s*<\/div>\s*)+$/gi, '')
      .trim();

    return result || html; // Fallback to original if stripping removed everything
  }

  // ============================================
  // renderEmailInIframe — Sandboxed iframe for email HTML
  // ============================================
  function renderEmailInIframe(html, container, isDark) {
    const temp = document.createElement('div');
    temp.innerHTML = html;
    temp.querySelectorAll('script, object, embed').forEach(el => el.remove());
    temp.querySelectorAll('*').forEach(el => {
      Array.from(el.attributes).forEach(attr => {
        if (attr.name.startsWith('on') || (attr.name === 'href' && attr.value.startsWith('javascript:'))) el.removeAttribute(attr.name);
      });
      // Strip min-width and fixed widths that would cause horizontal overflow
      const style = el.getAttribute('style') || '';
      if (style) {
        el.setAttribute('style', style
          .replace(/min-width\s*:\s*\d{3,}px/gi, 'min-width:0')
          .replace(/width\s*:\s*\d{3,}px/gi, 'width:100%')
        );
      }
      // Strip width/height HTML attributes on layout elements
      if (el.hasAttribute('width')) {
        const w = parseInt(el.getAttribute('width'));
        if (w > 100) el.removeAttribute('width');
      }
    });
    temp.querySelectorAll('a').forEach(link => { link.setAttribute('target', '_blank'); link.setAttribute('rel', 'noopener noreferrer'); });
    const cleanHtml = temp.innerHTML;

    const textColor = isDark ? '#e8e8e8' : '#2c3e50';
    const linkColor = isDark ? '#8b9ff0' : '#667eea';

    const iframeDoc = `<!DOCTYPE html>
<html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<style>
*{box-sizing:border-box;max-width:100%!important}
html{overflow:hidden}
body{margin:0;padding:8px 12px;background:transparent;color:${textColor};font-family:'Lato',-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;font-size:14px;line-height:1.6;word-wrap:break-word;overflow-wrap:break-word;overflow:hidden;height:auto;max-width:100%!important}
a{color:${linkColor};word-break:break-all}
img{max-width:100%!important;height:auto!important}
table{border-collapse:collapse;overflow:visible;max-width:100%!important;width:100%!important;table-layout:fixed}
col{width:auto!important}
td,th{word-break:normal;overflow-wrap:break-word;max-width:100%!important;width:auto!important}
span[style*="width"],div[style*="width"]{max-width:100%!important;overflow:hidden}
div,span,p,section,article{overflow-wrap:break-word;max-width:100%!important}
pre{white-space:pre-wrap;word-break:break-all;overflow-x:auto;max-width:100%}
blockquote{margin:8px 0;padding:8px 12px;border-left:3px solid ${linkColor};background:${isDark?'rgba(102,126,234,0.1)':'rgba(102,126,234,0.05)'};max-width:100%;overflow:auto}
hr{border:none;border-top:1px solid ${isDark?'rgba(255,255,255,0.12)':'rgba(0,0,0,0.08)'};margin:10px 0}
table[cellpadding],table[cellspacing]{font-size:13px}
h2{font-size:15px;margin:4px 0}p{margin:4px 0}
ul,ol{margin:4px 0;padding-left:20px}li{margin:2px 0;white-space:normal!important}
p,div,span,li,td,th{white-space:normal!important}
b,strong{font-weight:600;color:${textColor}}
${isDark ? `
/* Dark mode: override inline dark colors that become invisible */
body *:not(a){color:${textColor}!important}
a{color:${linkColor}!important}
/* Dark mode: override light backgrounds that clash */
body,table,tr,td,th,div,section,article,header,footer,main,aside,nav{background-color:transparent!important;background-image:none!important}
/* Preserve images but make container backgrounds transparent */
img{background-color:transparent!important}
/* Override common white/light background inline styles */
[style*="background:#fff"],[style*="background: #fff"],[style*="background-color:#fff"],[style*="background-color: #fff"],
[style*="background:#FFF"],[style*="background: #FFF"],[style*="background-color:#FFF"],[style*="background-color: #FFF"],
[style*="background:white"],[style*="background: white"],[style*="background-color:white"],[style*="background-color: white"],
[style*="background:#ffffff"],[style*="background: #ffffff"],[style*="background-color:#ffffff"],[style*="background-color: #ffffff"],
[style*="background:#FFFFFF"],[style*="background: #FFFFFF"],[style*="background-color:#FFFFFF"],[style*="background-color: #FFFFFF"],
[style*="background:#f"],[style*="background: #f"],[style*="background-color:#f"],[style*="background-color: #f"],
[style*="background:#e"],[style*="background: #e"],[style*="background-color:#e"],[style*="background-color: #e"],
[style*="background:#d"],[style*="background: #d"],[style*="background-color:#d"],[style*="background-color: #d"],
[style*="background:rgb(2"],[style*="background: rgb(2"],[style*="background-color:rgb(2"],[style*="background-color: rgb(2"]
{background-color:transparent!important;background-image:none!important}
/* Override border colors that are too light in dark mode */
td,th,table{border-color:rgba(255,255,255,0.1)!important}
hr{border-color:rgba(255,255,255,0.12)!important}
` : ''}
img[alt="mobilePhone"],img[alt="emailAddress"],img[alt="website"],img[alt="address"]{width:12px!important;height:auto}
div[class*="elementToProof"]{margin-top:0.4em;margin-bottom:0.4em}
</style></head><body>${cleanHtml}</body></html>`;

    const iframe = document.createElement('iframe');
    iframe.style.cssText = 'width:100%;border:none;min-height:60px;background:transparent;display:block;margin:0;padding:0;';
    iframe.setAttribute('sandbox', 'allow-same-origin allow-popups allow-popups-to-escape-sandbox');
    iframe.srcdoc = iframeDoc;

    iframe.onload = () => {
      try {
        const resize = () => {
          if (iframe.contentDocument?.body) {
            const bodyH = iframe.contentDocument.body.scrollHeight;
            const docH = iframe.contentDocument.documentElement.scrollHeight;
            const newHeight = Math.max(bodyH, docH) + 8;
            if (newHeight > 60) {
              iframe.style.height = newHeight + 'px';
            }
          }
        };
        // Hide iframe until fully sized to prevent incremental expansion
        iframe.style.visibility = 'hidden';
        iframe.style.transition = 'none';
        resize();
        // Final resize after images/fonts load, then reveal
        setTimeout(() => {
          resize();
          iframe.style.visibility = 'visible';
        }, 150);
        setTimeout(resize, 600);
        iframe.contentDocument.querySelectorAll('img').forEach(img => { if (!img.complete) img.addEventListener('load', resize); });
        // Use ResizeObserver on the iframe body for dynamic content
        if (typeof ResizeObserver !== 'undefined' && iframe.contentDocument?.body) {
          const ro = new ResizeObserver(resize);
          ro.observe(iframe.contentDocument.body);
        }
      } catch (e) { /* cross-origin iframe safety */ }
    };

    container.style.padding = '0';
    container.innerHTML = '';
    container.appendChild(iframe);
  }

  // ============================================
  // fetchGmailAttachment — Fetch via webhook
  // ============================================
  async function fetchGmailAttachment(messageId, attachmentId, mimeType, filename) {
    try {
      if (typeof showToastNotification === 'function') showToastNotification('Fetching image...');
      const response = await fetch(GMAIL_ATTACHMENT_WEBHOOK, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(createAuthenticatedPayload({
          action: 'fetch_gmail_attachment',
          messageId, attachmentId, mimeType, filename,
          timestamp: new Date().toISOString()
        }))
      });
      const data = await response.json();
      if (data.success && data.base64Data) {
        // Track which attachment was just fetched so gallery can match it
        window._lastFetchedGmailAttachmentId = attachmentId;
        showAttachmentPreview(`data:${mimeType};base64,${data.base64Data}`, mimeType, filename);
      } else {
        if (typeof showToastNotification === 'function') showToastNotification(data.error || 'Failed to fetch attachment');
      }
    } catch (e) {
      console.error('Error fetching Gmail attachment:', e);
      if (typeof showToastNotification === 'function') showToastNotification('Failed to fetch attachment');
    }
  }

  // ============================================
  // renderTranscriptAttachment — Build compact attachment element
  // ============================================
  function renderTranscriptAttachment(attachment) {
    const { name, type, url, size, text, attachmentId, messageId, isInline, contentId } = attachment;
    if (!name && !url && !text && !attachmentId) return null;

    const isDark = document.body.classList.contains('dark-mode');
    const attachEl = document.createElement('div');
    const isGmailAttachment = attachmentId && messageId;
    const lt = (type || '').toLowerCase();
    const ln = (name || '').toLowerCase();
    const lu = (url || '').toLowerCase();

    // Determine icon
    let icon = '📎', iconBg = 'linear-gradient(45deg,#667eea,#764ba2)';
    const isImage = lt.includes('image') || lt === 'jpeg' || lt === 'png' || lt === 'gif' ||
      ln.endsWith('.png') || ln.endsWith('.jpg') || ln.endsWith('.jpeg') || ln.endsWith('.gif') || ln.endsWith('.webp');
    const isPreviewable = isImage || lt.includes('video') || lt.includes('movie') || lt.includes('mpeg') ||
      lt.includes('pdf') || ln.endsWith('.mp4') || ln.endsWith('.mov') || ln.endsWith('.webm') || ln.endsWith('.pdf') ||
      lu.includes('.png') || lu.includes('.jpg') || lu.includes('.jpeg') || lu.includes('.mp4') || lu.includes('.pdf');

    const isLinkPreview = lt === 'link' || lt.includes('linkedin') || lt.includes('youtube') || lt.includes('google docs') ||
      lt.includes('google sheets') || lt.includes('google slides') || lt.includes('attio') || lt.includes('openai') ||
      lt.includes('freshworks') || (lt.includes('.com') || lt.includes('.in') || lt.includes('.io'));

    if (lt.includes('video') || lt.includes('movie') || lt.includes('mpeg') || ln.endsWith('.mov') || ln.endsWith('.mp4') || ln.endsWith('.avi') || ln.endsWith('.webm')) { icon = '🎬'; iconBg = 'linear-gradient(45deg,#e74c3c,#c0392b)'; }
    else if (lt.includes('youtube')) { icon = '▶️'; iconBg = 'linear-gradient(45deg,#ff0000,#cc0000)'; }
    else if (lt.includes('linkedin')) { icon = '💼'; iconBg = 'linear-gradient(45deg,#0077b5,#005885)'; }
    else if (lt.includes('audio') || ln.endsWith('.mp3') || ln.endsWith('.wav') || ln.endsWith('.m4a')) { icon = '🎵'; iconBg = 'linear-gradient(45deg,#9b59b6,#8e44ad)'; }
    else if (isImage) { icon = '🖼️'; iconBg = 'linear-gradient(45deg,#3498db,#2980b9)'; }
    else if (lt.includes('pdf') || ln.endsWith('.pdf')) { icon = '📄'; iconBg = 'linear-gradient(45deg,#e74c3c,#c0392b)'; }
    else if (lt.includes('google docs') || lt.includes('word') || lt.includes('document') || ln.endsWith('.doc') || ln.endsWith('.docx')) { icon = '📝'; iconBg = 'linear-gradient(45deg,#4285f4,#2a5db0)'; }
    else if (lt.includes('google sheets') || lt.includes('sheet') || lt.includes('excel') || ln.endsWith('.xls') || ln.endsWith('.xlsx') || ln.endsWith('.csv')) { icon = '📊'; iconBg = 'linear-gradient(45deg,#0f9d58,#0b7a45)'; }
    else if (lt.includes('google slides') || lt.includes('presentation') || lt.includes('powerpoint') || ln.endsWith('.ppt') || ln.endsWith('.pptx')) { icon = '📽️'; iconBg = 'linear-gradient(45deg,#f4b400,#c99200)'; }
    else if (lt.includes('zip') || lt.includes('archive') || ln.endsWith('.zip') || ln.endsWith('.rar') || ln.endsWith('.7z')) { icon = '📦'; iconBg = 'linear-gradient(45deg,#7f8c8d,#5d6d7e)'; }
    else if (isLinkPreview) { icon = '🔗'; iconBg = 'linear-gradient(45deg,#667eea,#764ba2)'; }

    let sizeStr = '';
    if (size) {
      if (size < 1024) sizeStr = size + ' B';
      else if (size < 1024 * 1024) sizeStr = (size / 1024).toFixed(1) + ' KB';
      else sizeStr = (size / (1024 * 1024)).toFixed(1) + ' MB';
    }

    let previewText = text && text.length > 0 ? (text.length > 150 ? text.substring(0, 150) + '...' : text) : '';
    const hasPreview = text && text.length > 0 && isLinkPreview;

    // Detect Slack message references: name like "[February 13th, 2026 8:30 AM] username: ..."
    const slackMsgRefMatch = (name || '').match(/^\[([^\]]+)\]\s+([^:]+):\s*(.*)/s);
    if (slackMsgRefMatch && isLinkPreview) {
      const refDate = slackMsgRefMatch[1];
      const refSender = slackMsgRefMatch[2].trim();
      const refSnippet = slackMsgRefMatch[3]?.trim() || '';
      // Use the text field for full content
      let fullContent = text || refSnippet || '';
      // Clean up Slack user ID mentions but keep text
      fullContent = fullContent.replace(/<@[A-Z0-9]+>/g, '');
      // Convert Slack links to readable format
      fullContent = fullContent.replace(/<(https?:\/\/[^|>]+)\|([^>]+)>/g, '$2 ($1)');
      fullContent = fullContent.replace(/<(https?:\/\/[^>]+)>/g, '$1');

      const isLong = fullContent.length > 200;
      const previewContent = isLong ? fullContent.substring(0, 200) + '…' : fullContent;

      const formattedPreview = typeof formatMessageContent === 'function' ? formatMessageContent(previewContent) : escapeHtml(previewContent);
      const formattedFull = isLong ? (typeof formatMessageContent === 'function' ? formatMessageContent(fullContent) : escapeHtml(fullContent)) : formattedPreview;

      const senderInitials = refSender.split(' ').map(n => n[0]).join('').substring(0, 2).toUpperCase();
      const senderColor = typeof getUserChipColor === 'function' ? getUserChipColor(refSender) : '#667eea';
      const uid = 'ref-' + Math.random().toString(36).substring(2, 8);

      attachEl.style.cssText = `display:flex;flex-direction:column;gap:0;padding:0;background:transparent;border-left:3px solid ${isDark ? 'rgba(102,126,234,0.5)' : 'rgba(102,126,234,0.4)'};border-radius:0 8px 8px 0;width:100%;overflow:hidden;margin:4px 0;`;
      attachEl.innerHTML = `
        <div style="padding:10px 12px;background:${isDark ? 'rgba(255,255,255,0.03)' : 'rgba(102,126,234,0.04)'};border:1px solid ${isDark ? 'rgba(255,255,255,0.06)' : 'rgba(225,232,237,0.6)'};border-left:none;border-radius:0 8px 8px 0;">
          <div style="display:flex;align-items:center;gap:8px;margin-bottom:6px;">
            <div style="width:22px;height:22px;background:${senderColor};border-radius:50%;display:flex;align-items:center;justify-content:center;color:white;font-size:9px;font-weight:600;flex-shrink:0;">${senderInitials}</div>
            <span style="font-weight:600;font-size:12px;color:${isDark ? '#ccc' : '#2c3e50'};">${escapeHtml(refSender)}</span>
            <span style="font-size:10px;color:${isDark ? '#666' : '#95a5a6'};">${escapeHtml(refDate)}</span>
          </div>
          <div id="${uid}-preview" style="font-size:12px;color:${isDark ? '#aaa' : '#5d6d7e'};line-height:1.5;overflow:hidden;word-break:break-word;">${formattedPreview}</div>
          ${isLong ? `<div id="${uid}-full" style="font-size:12px;color:${isDark ? '#aaa' : '#5d6d7e'};line-height:1.5;overflow:hidden;word-break:break-word;display:none;">${formattedFull}</div>` : ''}
          ${isLong ? `<button id="${uid}-toggle" style="background:none;border:none;color:#667eea;font-size:11px;font-weight:600;cursor:pointer;padding:4px 0 0 0;text-align:left;">View more</button>` : ''}
        </div>`;

      if (isLong) {
        setTimeout(() => {
          const toggle = document.getElementById(`${uid}-toggle`);
          const preview = document.getElementById(`${uid}-preview`);
          const full = document.getElementById(`${uid}-full`);
          if (toggle && preview && full) {
            let expanded = false;
            toggle.addEventListener('click', (e) => {
              e.stopPropagation();
              expanded = !expanded;
              preview.style.display = expanded ? 'none' : '';
              full.style.display = expanded ? '' : 'none';
              toggle.textContent = expanded ? 'View less' : 'View more';
            });
          }
        }, 0);
      }
      return attachEl;
    }

    // Link previews stay full-width
    if (hasPreview) {
      attachEl.style.cssText = `display:flex;flex-direction:column;gap:8px;padding:12px;background:${isDark ? 'rgba(255,255,255,0.04)' : 'white'};border:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.8)'};border-radius:10px;transition:all 0.2s;cursor:pointer;width:100%;`;
      attachEl.innerHTML = `
        <div style="display:flex;align-items:center;gap:12px;">
          <div style="width:36px;height:36px;background:${iconBg};border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0;">${icon}</div>
          <div style="flex:1;min-width:0;overflow:hidden;">
            <div style="font-weight:600;font-size:13px;color:${isDark ? '#e8e8e8' : '#2c3e50'};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${escapeHtml(name || 'Link')}">${escapeHtml(name || 'Link')}</div>
            <div style="font-size:11px;color:${isDark ? '#888' : '#7f8c8d'};">${type ? escapeHtml(type) : 'Link'}</div>
          </div>
          ${url ? `<a href="${escapeHtml(url)}" target="_blank" style="display:flex;align-items:center;justify-content:center;width:28px;height:28px;background:linear-gradient(45deg,#667eea,#764ba2);border-radius:6px;color:white;text-decoration:none;font-size:12px;flex-shrink:0;" title="Open link">↗</a>` : ''}
        </div>
        <div style="font-size:12px;color:${isDark ? '#aaa' : '#5d6d7e'};line-height:1.5;padding-left:48px;border-left:3px solid rgba(102,126,234,0.3);margin-left:16px;">${escapeHtml(previewText)}</div>`;
    } else {
      // Compact card layout — fits multiple per row
      const truncName = (name || 'Attachment').length > 18 ? (name || 'Attachment').substring(0, 15) + '...' : (name || 'Attachment');

      // For images with URL, show thumbnail; for Gmail attachments, show icon with fetch-on-click
      let thumbnailHtml;
      if (isImage && url) {
        thumbnailHtml = `<div class="attach-thumb" style="width:48px;height:48px;border-radius:8px;overflow:hidden;flex-shrink:0;background:${isDark ? 'rgba(255,255,255,0.05)' : '#f0f2f5'};"><img src="${escapeHtml(url)}" style="width:100%;height:100%;object-fit:cover;" data-fallback-icon="${icon}" data-fallback-bg="${iconBg}"></div>`;
      } else if (isImage && isGmailAttachment) {
        thumbnailHtml = `<div class="attach-thumb gmail-fetch-btn" data-message-id="${escapeHtml(messageId)}" data-attachment-id="${escapeHtml(attachmentId)}" data-mime-type="${escapeHtml(type || 'image/png')}" data-filename="${escapeHtml(name || 'attachment')}" style="width:48px;height:48px;border-radius:8px;overflow:hidden;flex-shrink:0;background:${iconBg};display:flex;align-items:center;justify-content:center;font-size:22px;cursor:pointer;" title="Click to load">${icon}</div>`;
      } else if (isGmailAttachment) {
        // Non-image Gmail attachment (docx, pdf, etc.) — add fetch button for download on click
        thumbnailHtml = `<div class="gmail-fetch-btn" data-message-id="${escapeHtml(messageId)}" data-attachment-id="${escapeHtml(attachmentId)}" data-mime-type="${escapeHtml(type || 'application/octet-stream')}" data-filename="${escapeHtml(name || 'attachment')}" style="width:40px;height:40px;background:${iconBg};border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0;cursor:pointer;" title="Click to download">${icon}</div>`;
      } else {
        thumbnailHtml = `<div style="width:40px;height:40px;background:${iconBg};border-radius:8px;display:flex;align-items:center;justify-content:center;font-size:18px;flex-shrink:0;">${icon}</div>`;
      }

      attachEl.style.cssText = `display:flex;align-items:center;gap:10px;padding:8px 10px;background:${isDark ? 'rgba(255,255,255,0.04)' : 'white'};border:1px solid ${isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.8)'};border-radius:10px;transition:all 0.2s;cursor:pointer;min-width:0;flex:1 1 auto;max-width:260px;`;

      attachEl.innerHTML = `
        ${thumbnailHtml}
        <div style="flex:1;min-width:0;overflow:hidden;">
          <div style="font-weight:600;font-size:12px;color:${isDark ? '#e8e8e8' : '#2c3e50'};white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${escapeHtml(name || 'Attachment')}">${escapeHtml(truncName)}</div>
          <div style="font-size:10px;color:${isDark ? '#888' : '#7f8c8d'};margin-top:1px;">${sizeStr || (type ? escapeHtml(type) : '')}</div>
        </div>`;
    }

    // Image thumbnail fallback (CSP-safe, no inline onerror)
    const thumbImg = attachEl.querySelector('.attach-thumb img[data-fallback-icon]');
    if (thumbImg) {
      thumbImg.addEventListener('error', function() {
        const icon = this.dataset.fallbackIcon;
        const bg = this.dataset.fallbackBg;
        this.parentElement.innerHTML = `<div style="width:100%;height:100%;background:${bg};display:flex;align-items:center;justify-content:center;font-size:22px;">${icon}</div>`;
      });
    }

    // Hover effects
    attachEl.addEventListener('mouseenter', () => { attachEl.style.background = isDark ? 'rgba(102,126,234,0.1)' : 'rgba(102,126,234,0.05)'; attachEl.style.borderColor = 'rgba(102,126,234,0.3)'; });
    attachEl.addEventListener('mouseleave', () => { attachEl.style.background = isDark ? 'rgba(255,255,255,0.04)' : 'white'; attachEl.style.borderColor = isDark ? 'rgba(255,255,255,0.1)' : 'rgba(225,232,237,0.8)'; });

    // Click handlers
    if (isGmailAttachment) {
      const gmailBtn = attachEl.querySelector('.gmail-fetch-btn');
      const fetchGmail = (el) => {
        fetchGmailAttachment(el.getAttribute('data-message-id'), el.getAttribute('data-attachment-id'), el.getAttribute('data-mime-type'), el.getAttribute('data-filename'));
      };
      if (gmailBtn) gmailBtn.addEventListener('click', (e) => { e.stopPropagation(); fetchGmail(gmailBtn); });
      attachEl.addEventListener('click', (e) => { if (e.target.closest('.gmail-fetch-btn')) return; if (gmailBtn) fetchGmail(gmailBtn); });
    } else if (url) {
      attachEl.addEventListener('click', (e) => { if (e.target.closest('a')) return; isPreviewable ? showAttachmentPreview(url, type, name) : window.open(url, '_blank'); });
    }

    return attachEl;
  }

  // ============================================
  // isSlackPrivateUrl — Detect Slack private file URLs
  // ============================================
  function isSlackPrivateUrl(url) {
    return url && (url.includes('files.slack.com/files-pri/') || url.includes('files.slack.com/files-tmb/'));
  }

  // ============================================
  // ============================================
  // collectPreviewableAttachments — Gather all previewable images in any open slider/transcript
  // Returns array of { url, type, name } for gallery navigation
  // ============================================
  function collectPreviewableAttachments() {
    const items = [];
    const seen = new Set();

    // Find all image thumbnails inside attachment cards anywhere in the DOM
    document.querySelectorAll('.attach-thumb img').forEach(img => {
      const src = img.getAttribute('src');
      if (src && !seen.has(src)) {
        seen.add(src);
        const parentCard = img.closest('[style*="cursor"]');
        const nameEl = parentCard?.querySelector('[title]');
        const name = nameEl?.getAttribute('title') || img.getAttribute('alt') || 'Image';
        items.push({ url: src, type: 'image', name });
      }
    });

    // Collect data:image URLs from transcript areas
    document.querySelectorAll('.transcript-messages-container img[src^="data:image"], .transcript-message-attachments img[src^="data:image"]').forEach(img => {
      const src = img.getAttribute('src');
      if (src && !seen.has(src)) {
        seen.add(src);
        items.push({ url: src, type: 'image', name: img.getAttribute('alt') || 'Image' });
      }
    });

    // Also include unfetched Gmail image attachments (so gallery knows about them)
    document.querySelectorAll('.gmail-fetch-btn[data-mime-type*="image"]').forEach(btn => {
      const id = btn.dataset.attachmentId;
      if (id && !seen.has('gmail:' + id)) {
        seen.add('gmail:' + id);
        items.push({
          url: 'gmail:' + id,
          type: btn.dataset.mimeType || 'image',
          name: btn.dataset.filename || 'Image',
          _gmailFetch: { messageId: btn.dataset.messageId, attachmentId: id, mimeType: btn.dataset.mimeType, filename: btn.dataset.filename }
        });
      }
    });

    return items;
  }

  // showAttachmentPreview — Full-screen modal for images/videos/PDFs
  // Now supports gallery navigation for multiple images
  // ============================================
  function showAttachmentPreview(url, type, name) {
    document.querySelectorAll('.attachment-preview-modal').forEach(m => m.remove());
    const lt = (type || '').toLowerCase(), ln = (name || '').toLowerCase(), lu = (url || '').toLowerCase();

    // Slack private file URLs (non-image) can't be previewed inline — open in new tab
    const isImageFile = lt.includes('image') || lt === 'jpeg' || lt === 'png' || lt === 'gif' ||
      ln.endsWith('.png') || ln.endsWith('.jpg') || ln.endsWith('.jpeg') || ln.endsWith('.gif') || ln.endsWith('.webp');
    if (isSlackPrivateUrl(url) && !isImageFile) {
      window.open(url, '_blank');
      return;
    }

    let contentType = 'unknown';
    if (lt.includes('image') || lt === 'jpeg' || lt === 'png' || lt === 'gif' || ln.endsWith('.png') || ln.endsWith('.jpg') || ln.endsWith('.jpeg') || ln.endsWith('.gif') || ln.endsWith('.webp') || lu.includes('.png') || lu.includes('.jpg') || lu.includes('.jpeg') || lu.startsWith('data:image')) contentType = 'image';
    else if (lt.includes('video') || lt.includes('movie') || ln.endsWith('.mp4') || ln.endsWith('.mov') || ln.endsWith('.webm') || lu.includes('.mp4') || lu.includes('.mov')) contentType = 'video';
    else if (lt.includes('pdf') || ln.endsWith('.pdf') || lu.includes('.pdf')) contentType = 'pdf';

    if (contentType === 'unknown') { window.open(url, '_blank'); return; }

    // Collect all previewable images for gallery navigation
    let gallery = [];
    let currentGalleryIndex = -1;
    if (contentType === 'image') {
      gallery = collectPreviewableAttachments();
      currentGalleryIndex = gallery.findIndex(item => item.url === url);
      // For data: URLs from Gmail fetch, match by tracked attachment ID
      if (currentGalleryIndex === -1 && url.startsWith('data:')) {
        const lastFetchedId = window._lastFetchedGmailAttachmentId;
        if (lastFetchedId) {
          currentGalleryIndex = gallery.findIndex(item => item.url === 'gmail:' + lastFetchedId);
        }
        // Fallback: match by name
        if (currentGalleryIndex === -1) {
          currentGalleryIndex = gallery.findIndex(item => item.url.startsWith('gmail:') && item.name === name);
        }
        if (currentGalleryIndex !== -1) {
          // Replace the placeholder with the actual fetched data URL
          gallery[currentGalleryIndex].url = url;
          gallery[currentGalleryIndex]._gmailFetch = null;
        }
        window._lastFetchedGmailAttachmentId = null;
      }
      // If still not found, insert at position 0 and look for other gallery items
      if (currentGalleryIndex === -1) {
        if (gallery.length > 0) {
          gallery.unshift({ url, type, name });
          currentGalleryIndex = 0;
        } else {
          gallery = [{ url, type, name }];
          currentGalleryIndex = 0;
        }
      }
    }
    const hasGallery = gallery.length > 1;

    const modal = document.createElement('div');
    modal.className = 'attachment-preview-modal';
    modal.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.9);z-index:100000;display:flex;flex-direction:column;animation:fadeIn 0.2s ease-out;';

    const header = document.createElement('div');
    header.style.cssText = 'display:flex;align-items:center;justify-content:space-between;padding:16px 20px;background:rgba(0,0,0,0.5);flex-shrink:0;';

    const headerLeft = document.createElement('div');
    headerLeft.style.cssText = 'display:flex;align-items:center;gap:12px;color:white;min-width:0;flex:1;';
    headerLeft.innerHTML = `
      <span style="font-size:20px;">${contentType === 'image' ? '🖼️' : contentType === 'video' ? '🎬' : '📄'}</span>
      <span class="preview-title" style="font-weight:600;font-size:14px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;" title="${escapeHtml(name || 'Preview')}">${escapeHtml(name || 'Preview')}</span>
      ${hasGallery ? `<span class="preview-counter" style="font-size:12px;color:rgba(255,255,255,0.6);flex-shrink:0;">${currentGalleryIndex + 1} / ${gallery.length}</span>` : ''}`;

    const headerRight = document.createElement('div');
    headerRight.style.cssText = 'display:flex;align-items:center;gap:8px;flex-shrink:0;';
    headerRight.innerHTML = `
      <a class="preview-open-link" href="${escapeHtml(url)}" target="_blank" style="display:flex;align-items:center;gap:6px;padding:8px 16px;background:rgba(255,255,255,0.1);border:1px solid rgba(255,255,255,0.2);border-radius:8px;color:white;text-decoration:none;font-size:13px;font-weight:500;"><span>↗</span> Open in new tab</a>
      <button class="preview-close-btn" style="display:flex;align-items:center;justify-content:center;width:40px;height:40px;background:rgba(231,76,60,0.8);border:none;border-radius:8px;color:white;font-size:20px;cursor:pointer;">×</button>`;

    header.appendChild(headerLeft);
    header.appendChild(headerRight);

    const content = document.createElement('div');
    content.style.cssText = 'flex:1;display:flex;align-items:center;justify-content:center;padding:20px;overflow:auto;min-height:0;position:relative;';

    function renderContent(u, t, n) {
      if (t === 'image') return `<img src="${escapeHtml(u)}" alt="${escapeHtml(n || 'Image')}" style="max-width:100%;max-height:100%;object-fit:contain;border-radius:8px;box-shadow:0 8px 32px rgba(0,0,0,0.3);transition:opacity 0.2s;" />`;
      if (t === 'video') return `<video controls autoplay style="max-width:100%;max-height:100%;border-radius:8px;box-shadow:0 8px 32px rgba(0,0,0,0.3);"><source src="${escapeHtml(u)}" type="video/mp4">Your browser does not support video playback.</video>`;
      if (t === 'pdf') {
        const gv = `https://docs.google.com/viewer?url=${encodeURIComponent(u)}&embedded=true`;
        return `<div style="width:100%;height:100%;display:flex;flex-direction:column;gap:12px;"><object data="${escapeHtml(u)}" type="application/pdf" style="width:100%;height:100%;border-radius:8px;background:white;"><iframe src="${gv}" style="width:100%;height:100%;border:none;border-radius:8px;background:white;"></iframe></object></div>`;
      }
      return '';
    }

    content.innerHTML = renderContent(url, contentType, name);
    if (contentType === 'pdf') content.style.padding = '20px';

    // Navigation arrows for gallery
    const navBtnStyle = 'position:absolute;top:50%;transform:translateY(-50%);width:48px;height:48px;background:rgba(255,255,255,0.15);border:1px solid rgba(255,255,255,0.25);border-radius:50%;color:white;font-size:22px;cursor:pointer;display:flex;align-items:center;justify-content:center;transition:all 0.2s;z-index:2;backdrop-filter:blur(4px);';

    if (hasGallery) {
      const prevBtn = document.createElement('button');
      prevBtn.className = 'preview-nav-prev';
      prevBtn.style.cssText = navBtnStyle + 'left:16px;';
      prevBtn.innerHTML = '‹';
      prevBtn.title = 'Previous image (←)';
      prevBtn.addEventListener('mouseenter', () => { prevBtn.style.background = 'rgba(255,255,255,0.3)'; });
      prevBtn.addEventListener('mouseleave', () => { prevBtn.style.background = 'rgba(255,255,255,0.15)'; });

      const nextBtn = document.createElement('button');
      nextBtn.className = 'preview-nav-next';
      nextBtn.style.cssText = navBtnStyle + 'right:16px;';
      nextBtn.innerHTML = '›';
      nextBtn.title = 'Next image (→)';
      nextBtn.addEventListener('mouseenter', () => { nextBtn.style.background = 'rgba(255,255,255,0.3)'; });
      nextBtn.addEventListener('mouseleave', () => { nextBtn.style.background = 'rgba(255,255,255,0.15)'; });

      function navigateGallery(direction) {
        currentGalleryIndex = (currentGalleryIndex + direction + gallery.length) % gallery.length;
        const item = gallery[currentGalleryIndex];
        const imgEl = content.querySelector('img');
        const titleEl = headerLeft.querySelector('.preview-title');
        const counterEl = headerLeft.querySelector('.preview-counter');
        const openLink = headerRight.querySelector('.preview-open-link');

        // Update counter immediately
        if (counterEl) counterEl.textContent = `${currentGalleryIndex + 1} / ${gallery.length}`;
        if (titleEl) { titleEl.textContent = item.name || 'Preview'; titleEl.title = item.name || 'Preview'; }

        // If this is an unfetched Gmail attachment, fetch it
        if (item._gmailFetch && item.url.startsWith('gmail:')) {
          if (imgEl) { imgEl.style.opacity = '0.3'; }
          fetchGmailAttachment(item._gmailFetch.messageId, item._gmailFetch.attachmentId, item._gmailFetch.mimeType, item._gmailFetch.filename)
            .then(() => {
              // After fetch, the modal will be replaced — but we can update the gallery entry
              // The fetchGmailAttachment opens a new modal, so we don't need to do more here
            });
          return;
        }

        // Regular image - swap with fade
        if (imgEl) {
          imgEl.style.opacity = '0';
          setTimeout(() => {
            imgEl.src = item.url;
            imgEl.alt = item.name || 'Image';
            imgEl.style.opacity = '1';
          }, 150);
        }
        if (openLink) openLink.href = item.url;
      }

      prevBtn.addEventListener('click', (e) => { e.stopPropagation(); navigateGallery(-1); });
      nextBtn.addEventListener('click', (e) => { e.stopPropagation(); navigateGallery(1); });
      content.appendChild(prevBtn);
      content.appendChild(nextBtn);
    }

    modal.appendChild(header);
    modal.appendChild(content);
    document.body.appendChild(modal);

    const closeModal = () => { modal.style.animation = 'fadeOut 0.2s ease-out'; document.removeEventListener('keydown', keyH); setTimeout(() => modal.remove(), 180); };
    modal.querySelector('.preview-close-btn').addEventListener('click', (e) => { e.stopPropagation(); closeModal(); });
    modal.addEventListener('click', (e) => { e.stopPropagation(); if (e.target === modal) closeModal(); });
    content.addEventListener('click', (e) => { e.stopPropagation(); if (e.target === content) closeModal(); });
    const keyH = (e) => {
      if (e.key === 'Escape') { e.stopImmediatePropagation(); e.preventDefault(); closeModal(); }
      else if (hasGallery && e.key === 'ArrowLeft') { e.preventDefault(); content.querySelector('.preview-nav-prev')?.click(); }
      else if (hasGallery && e.key === 'ArrowRight') { e.preventDefault(); content.querySelector('.preview-nav-next')?.click(); }
    };
    document.addEventListener('keydown', keyH);
  }

  // ============================================
  // EXPORT
  // ============================================
  window.OracleMessageFormat = {
    formatMessageContent,
    sanitizeHtml,
    isComplexEmailHtml,
    stripEmailQuotedContent,
    renderEmailInIframe,
    fetchGmailAttachment,
    isSlackPrivateUrl,
    renderTranscriptAttachment,
    showAttachmentPreview,
  };

})();
