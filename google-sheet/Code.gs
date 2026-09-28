// The Hidden Brain — Google Sheet collector + designed email
// Paste into Extensions → Apps Script, save, then Deploy → Manage deployments → Edit → Version: New version → Deploy.

const SHEET_NAME = 'Submissions';
const SEND_EMAIL = true;          // email the blueprint to the user
const ATTACH_FILES = true;        // attach blueprint.md + prompts.md
const NOTIFY_ME = '';             // optional: your email for a copy of every signup
const FROM_NAME = 'The Hidden Brain';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents || '{}');
    const email = String(data.email || '').trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) return json({ ok: false, error: 'invalid email' });

    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(SHEET_NAME) || ss.insertSheet(SHEET_NAME);
      if (sh.getLastRow() === 0) sh.appendRow(['Date', 'Email', 'Page', 'Blueprint']);
      sh.appendRow([new Date(), email, data.page || '', String(data.blueprint || '').slice(0, 45000)]);
    } finally { lock.releaseLock(); }

    if (SEND_EMAIL) {
      const opts = { to: email, name: FROM_NAME, subject: 'Your AI Work Brain Blueprint', body: String(data.blueprint || ''), htmlBody: renderEmail(data) };
      if (ATTACH_FILES) opts.attachments = [
        Utilities.newBlob(String(data.blueprint || ''), 'text/markdown', 'ai-work-brain-blueprint.md'),
        Utilities.newBlob(promptsMd(data), 'text/markdown', 'ai-work-brain-prompts.md'),
      ];
      MailApp.sendEmail(opts);
    }
    if (NOTIFY_ME) MailApp.sendEmail(NOTIFY_ME, 'New Hidden Brain signup: ' + email, String(data.blueprint || ''));
    return json({ ok: true });
  } catch (err) {
    return json({ ok: false, error: String(err) });
  }
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function esc(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function promptsMd(d) {
  return '# AI Work Brain — Prompt Toolkit\n\n' + (d.prompts || []).map(function (p, i) {
    return '## 0' + (i + 1) + ' ' + p.n + '\n_' + p.t + '_\n\n' + p.d + '\n\n```\n' + p.body + '\n```';
  }).join('\n\n---\n\n') + '\n';
}

// ---------- designed email (table layout, inline styles for Gmail/Outlook) ----------
const C = { bg: '#05070c', card: '#0b1019', line: '#1c2536', ink: '#eef2fa', body: '#c3cbdc', mute: '#7d89a3', accent: '#c9d8ff' };
const SERIF = "'Instrument Serif',Georgia,'Times New Roman',serif";
const MONO = "'Geist Mono',ui-monospace,Menlo,Consolas,monospace";
const SANS = "-apple-system,BlinkMacSystemFont,'Segoe UI',Helvetica,Arial,sans-serif";

function label(t) {
  return '<div style="font-family:' + MONO + ';font-size:11px;letter-spacing:3px;text-transform:uppercase;color:' + C.mute + ';">' + esc(t) + '</div>';
}

function sectionBlock(s, i) {
  const lines = (s.lines || []).map(function (l) {
    const m = String(l).match(/^([^:—]{1,40}?)(: | — )([\s\S]*)$/);
    const inner = m
      ? '<span style="color:' + C.ink + ';font-weight:600;">' + esc(m[1]) + '</span><span style="color:' + C.mute + ';">' + esc(m[2]) + '</span>' + esc(m[3])
      : esc(l);
    return '<tr><td width="18" valign="top" style="padding:6px 0 0;"><div style="width:6px;height:6px;border-radius:3px;background:' + C.accent + ';margin-top:5px;"></div></td>' +
      '<td style="padding:4px 0;font-family:' + SANS + ';font-size:15px;line-height:23px;color:' + C.body + ';">' + inner + '</td></tr>';
  }).join('');
  return '<tr><td style="padding:0 0 14px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + C.card + ';border:1px solid ' + C.line + ';border-radius:16px;">' +
    '<tr><td style="padding:24px 26px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' +
    '<td valign="top" width="44" style="font-family:' + MONO + ';font-size:12px;letter-spacing:2px;color:' + C.mute + ';padding-top:8px;">0' + (i + 1) + '</td>' +
    '<td valign="top">' +
    '<div style="font-family:' + SERIF + ';font-size:26px;line-height:30px;color:' + C.ink + ';">' + esc(s.t) + '</div>' +
    '<div style="font-family:' + SANS + ';font-size:13px;line-height:20px;color:' + C.mute + ';padding:4px 0 12px;">' + esc(s.d) + '</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' + lines + '</table>' +
    '</td></tr></table></td></tr></table></td></tr>';
}

function renderEmail(d) {
  const sections = (d.sections && d.sections.length) ? d.sections : [{ t: 'Your Blueprint', d: '', lines: String(d.blueprint || '').split('\n').filter(Boolean) }];
  const prompts = d.prompts || [];
  const a = d.author || {};
  const toolkit = prompts.map(function (p, i) {
    return '<tr><td style="padding:12px 0;border-top:1px solid ' + C.line + ';">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' +
      '<td width="44" valign="top" style="font-family:' + MONO + ';font-size:12px;color:' + C.mute + ';padding-top:3px;">0' + (i + 1) + '</td>' +
      '<td valign="top"><div style="font-family:' + SANS + ';font-size:15px;font-weight:600;color:' + C.ink + ';">' + esc(p.n) + '</div>' +
      '<div style="font-family:' + SANS + ';font-size:13px;line-height:20px;color:' + C.mute + ';padding-top:2px;">' + esc(p.t) + '</div></td>' +
      '</tr></table></td></tr>';
  }).join('');
  const exp = prompts.length ? prompts[prompts.length - 1] : null;

  return '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><meta name="color-scheme" content="dark"><meta name="supported-color-schemes" content="dark">' +
    '<link href="https://fonts.googleapis.com/css2?family=Instrument+Serif&family=Geist+Mono&display=swap" rel="stylesheet"></head>' +
    '<body style="margin:0;padding:0;background:' + C.bg + ';">' +
    '<div style="display:none;max-height:0;overflow:hidden;opacity:0;">Your AI Work Brain Blueprint — seven decisions and the prompts to build it.</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:' + C.bg + ';"><tr><td align="center" style="padding:40px 16px;">' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="max-width:620px;">' +

    // header
    '<tr><td style="padding:0 4px 36px;">' + label('The Hidden Brain') +
    '<div style="font-family:' + SERIF + ';font-size:44px;line-height:48px;color:' + C.ink + ';padding:18px 0 12px;">Your AI Work Brain Blueprint</div>' +
    '<div style="font-family:' + SERIF + ';font-style:italic;font-size:20px;line-height:28px;color:' + C.body + ';">What do you know that your tools don\u2019t?</div>' +
    '<div style="height:1px;background:' + C.line + ';margin-top:28px;"></div></td></tr>' +

    // blueprint
    '<tr><td style="padding:0 4px 16px;">' + label('Your blueprint · 7 decisions') + '</td></tr>' +
    sections.map(sectionBlock).join('') +

    // first experiment prompt
    (exp ? '<tr><td style="padding:26px 0 14px;">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="border:1px solid #3a4a6b;border-radius:16px;background:#0d1422;"><tr><td style="padding:26px;">' +
      label('Start here') +
      '<div style="font-family:' + SERIF + ';font-size:28px;line-height:32px;color:' + C.ink + ';padding:12px 0 6px;">' + esc(exp.n) + '</div>' +
      '<div style="font-family:' + SANS + ';font-size:14px;line-height:21px;color:' + C.body + ';padding-bottom:16px;">' + esc(exp.d) + ' Paste it into your AI assistant.</div>' +
      '<div style="font-family:' + MONO + ';font-size:12.5px;line-height:20px;color:' + C.body + ';background:' + C.bg + ';border:1px solid ' + C.line + ';border-radius:10px;padding:16px 18px;white-space:pre-wrap;word-break:break-word;">' + esc(exp.body) + '</div>' +
      '</td></tr></table></td></tr>' : '') +

    // toolkit
    (prompts.length ? '<tr><td style="padding:26px 4px 8px;">' + label('Your prompt toolkit') +
      '<div style="font-family:' + SANS + ';font-size:14px;line-height:21px;color:' + C.mute + ';padding:10px 0 8px;">All seven prompts, pre-filled with your answers, are attached as <span style="color:' + C.ink + ';">ai-work-brain-prompts.md</span>.</div>' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' + toolkit + '</table></td></tr>' : '') +

    // cta
    (d.page ? '<tr><td align="left" style="padding:30px 4px 10px;"><a href="' + esc(d.page) + '" style="display:inline-block;padding:14px 26px;border-radius:999px;border:1px solid #8ea3cc;color:' + C.ink + ';text-decoration:none;font-family:' + MONO + ';font-size:12px;letter-spacing:3px;text-transform:uppercase;">Return to your brain \u2192</a></td></tr>' : '') +

    // footer
    '<tr><td style="padding:40px 4px 0;"><div style="height:1px;background:' + C.line + ';margin-bottom:20px;"></div>' +
    '<div style="font-family:' + SANS + ';font-size:12px;line-height:19px;color:' + C.mute + ';">\u00a9 2026 ' + esc(a.name || 'The Hidden Brain') +
    (a.linkedin ? ' · <a href="' + esc(a.linkedin) + '" style="color:' + C.accent + ';text-decoration:none;">LinkedIn</a>' : '') +
    (a.email ? ' · <a href="mailto:' + esc(a.email) + '" style="color:' + C.accent + ';text-decoration:none;">' + esc(a.email) + '</a>' : '') +
    '<br>You received this because you asked for your blueprint. One email, no list.</div></td></tr>' +

    '</table></td></tr></table></body></html>';
}

// Run this from the editor to preview the email in your own inbox.
function testEmail() {
  const me = NOTIFY_ME || Session.getActiveUser().getEmail();
  Logger.log('Sending test to: ' + me + ' · daily quota left: ' + MailApp.getRemainingDailyQuota());
  const d = {
    page: 'https://example.com',
    blueprint: '# Test',
    sections: [{ t: 'My Opportunity', d: 'The recurring problem I want to improve.', lines: ['Weekly client reporting', 'Answering the same onboarding questions'] },
               { t: 'My Boundaries', d: 'What AI can do, cannot do, and where people stay involved.', lines: ['Can do — draft summaries; tag requests', 'Human first — pricing changes'] }],
    prompts: [{ n: 'First Experiment Builder', t: 'Prototype the smallest useful version.', d: 'Everything you chose, assembled into one brief.', body: 'You are helping me design a small AI experiment…' }],
    author: { name: 'Your Name' },
  };
  MailApp.sendEmail({ to: me, subject: '[Test] Your AI Work Brain Blueprint', body: 'test', htmlBody: renderEmail(d), attachments: [Utilities.newBlob(promptsMd(d), 'text/markdown', 'ai-work-brain-prompts.md')] });
}
