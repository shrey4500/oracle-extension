// oracle-icons.js — Link detection, icon mapping, secondary links builder
// Used by both newtab.js and sidepanel.js

(function () {
  'use strict';

  const MEETING_LINK_PATTERNS = [
    'zoom.us', 'zoom.com', 'meet.google.com', 'teams.microsoft.com',
    'teams.live.com', 'webex.com', 'gotomeeting.com', 'bluejeans.com',
    'calendar.google.com', 'google.com/calendar'
  ];

  const DRIVE_LINK_PATTERNS = [
    'docs.google.com/document', 'docs.google.com/spreadsheets',
    'docs.google.com/presentation', 'docs.google.com/forms',
    'drive.google.com', 'sheets.google.com', 'slides.google.com'
  ];

  function isMeetingLink(url) {
    if (!url) return false;
    return MEETING_LINK_PATTERNS.some(p => url.includes(p));
  }

  function isDriveLink(url) {
    if (!url) return false;
    return DRIVE_LINK_PATTERNS.some(p => url.includes(p));
  }

  function isSlackLink(url) {
    if (!url) return false;
    return url.includes('slack.com') || url.includes('app.slack.com');
  }

  function getSlackChannelUrl(messageLink) {
    if (!messageLink || !isSlackLink(messageLink)) return '';
    const match = messageLink.match(/(https?:\/\/[^/]+\.slack\.com\/archives\/[A-Z0-9]+)/i);
    return match ? match[1] : messageLink;
  }

  function extractDriveFileId(url) {
    if (!url) return null;
    const patterns = [
      /docs\.google\.com\/document\/d\/([a-zA-Z0-9_-]+)/,
      /docs\.google\.com\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/,
      /docs\.google\.com\/presentation\/d\/([a-zA-Z0-9_-]+)/,
      /docs\.google\.com\/forms\/d\/([a-zA-Z0-9_-]+)/,
      /drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/,
      /drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/,
      /sheets\.google\.com\/.*\/d\/([a-zA-Z0-9_-]+)/,
      /slides\.google\.com\/.*\/d\/([a-zA-Z0-9_-]+)/,
      /\/d\/([a-zA-Z0-9_-]+)/  // fallback short pattern
    ];
    for (const pattern of patterns) {
      const match = url.match(pattern);
      if (match && match[1]) return match[1];
    }
    return null;
  }

  function getCleanDriveFileUrl(url, fileId) {
    if (!url || !fileId) return url;
    if (url.includes('docs.google.com/document')) return `https://docs.google.com/document/d/${fileId}/edit`;
    if (url.includes('docs.google.com/spreadsheets') || url.includes('sheets.google.com')) return `https://docs.google.com/spreadsheets/d/${fileId}/edit`;
    if (url.includes('docs.google.com/presentation') || url.includes('slides.google.com')) return `https://docs.google.com/presentation/d/${fileId}/edit`;
    if (url.includes('docs.google.com/forms')) return `https://docs.google.com/forms/d/${fileId}/edit`;
    if (url.includes('drive.google.com')) return `https://drive.google.com/file/d/${fileId}/view`;
    return url;
  }

  // Returns { icon: string (HTML), title: string }
  function getIconForLink(url) {
    if (!url) return { icon: '🔗', title: 'View link' };
    try {
      const img = (src, alt) => `<img src="${chrome.runtime.getURL(src)}" alt="${alt}" style="width:14px;height:14px;object-fit:contain;">`;
      if (url.includes('zoom.us') || url.includes('zoom.com')) return { icon: img('icon-zoom.png', 'Zoom'), title: 'Open in Zoom' };
      if (url.includes('meet.google.com')) return { icon: img('icon-google-meet.png', 'Meet'), title: 'Open in Google Meet' };
      if (url.includes('calendar.google.com') || url.includes('google.com/calendar')) return { icon: img('icon-google-calendar.png', 'Calendar'), title: 'Open in Google Calendar' };
      if (url.includes('mail.google.com')) return { icon: img('icon-gmail.png', 'Gmail'), title: 'Open in Gmail' };
      if (url.includes('slack.com')) return { icon: img('icon-slack.png', 'Slack'), title: 'Open in Slack' };
      if (url.includes('freshdesk.com')) return { icon: img('icon-freshdesk.png', 'Freshdesk'), title: 'Open in Freshdesk' };
      if (url.includes('freshrelease.com')) return { icon: img('icon-freshrelease.png', 'Freshrelease'), title: 'Open in Freshrelease' };
      if (url.includes('freshservice.com')) return { icon: img('icon-freshservice.png', 'Freshservice'), title: 'Open in Freshservice' };
      if (url.includes('docs.google.com/document')) return { icon: img('icon-google-docs.png', 'Docs'), title: 'Open in Google Docs' };
      if (url.includes('docs.google.com/spreadsheets') || url.includes('sheets.google.com')) return { icon: img('icon-google-sheets.png', 'Sheets'), title: 'Open in Google Sheets' };
      if (url.includes('docs.google.com/presentation') || url.includes('slides.google.com')) return { icon: img('icon-google-slides.png', 'Slides'), title: 'Open in Google Slides' };
      if (url.includes('drive.google.com')) return { icon: img('icon-drive.png', 'Drive'), title: 'Open in Google Drive' };
    } catch (e) { /* extension URL may fail in some contexts */ }
    return { icon: '🔗', title: 'View source' };
  }

  // Build secondary links HTML for todo items
  function buildSecondaryLinksHtml(secondaryLinks) {
    if (!secondaryLinks || !secondaryLinks.length) return '';
    return secondaryLinks.map(link => {
      const iconData = getIconForLink(link);
      return `<span style="color:#bdc3c7;margin:0 4px;">|</span><a href="${link}" target="_blank" class="todo-source" title="${iconData.title}" style="display:inline-flex;align-items:center;justify-content:center;width:22px;height:22px;border-radius:4px;background:rgba(102,126,234,0.08);transition:all 0.2s;">${iconData.icon}</a>`;
    }).join('');
  }

  // ============================================
  // EXPORT
  // ============================================
  window.OracleIcons = {
    MEETING_LINK_PATTERNS,
    DRIVE_LINK_PATTERNS,
    isMeetingLink,
    isDriveLink,
    isSlackLink,
    getSlackChannelUrl,
    extractDriveFileId,
    getCleanDriveFileUrl,
    getIconForLink,
    buildSecondaryLinksHtml,
  };

})();
