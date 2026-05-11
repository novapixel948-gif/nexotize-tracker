// ============================================================
// FULL OUTBOUND COLD EMAIL SYSTEM - Google Apps Script
// Sheet Format: Instagram Leads (your existing sheet)
// ============================================================

const CONFIG = {
  SHEET_ID:        "1-kswXjKtkjUa10tT-o2p-D5qEA7Nh0WSsNhHI5x1ki4",
  OPENAI_API_KEY:  "YOUR_NEW_OPENAI_KEY",       // rotate karo — purani expose ho gayi
  TRACKING_SERVER: "https://placeholder.com",   // Railway deploy ke baad update karna

  SENDER_ACCOUNTS: [
    { email: "contact.elenabrooks@gmail.com", alias: "Alex from Nexotize" },
    // Baad mein add karna:
    // { email: "outreach2@gmail.com", alias: "Alex | Nexotize Media" },
  ],

  DAILY_LIMIT_PER_ACCOUNT: 40,
  MIN_DELAY_MS: 60000,
  MAX_DELAY_MS: 300000,
};

// ─── YOUR INSTAGRAM SHEET COLUMNS ───────────────────────────
const COL = {
  INSTAGRAM_ID:   1,  // A
  USERNAME:       2,  // B
  FULL_NAME:      3,  // C
  PROFILE_LINK:   4,  // D
  AVATAR:         5,  // E
  FOLLOWED_BY:    6,  // F
  IS_VERIFIED:    7,  // G
  FOLLOWERS:      8,  // H
  FOLLOWING:      9,  // I
  BIOGRAPHY:      10, // J
  CATEGORY:       11, // K
  PUBLIC_EMAIL:   12, // L
  POSTS_COUNT:    13, // M
  PHONE_COUNTRY:  14, // N
  PHONE_NUMBER:   15, // O
  CITY:           16, // P
  ADDRESS:        17, // Q
  IS_PRIVATE:     18, // R
  IS_BUSINESS:    19, // S
  EXTERNAL_URL:   20, // T

  // New columns added at the end (U onwards)
  STATUS:         21, // U
  SEQ_DAY:        22, // V
  LAST_SENT:      23, // W
  OPENS:          24, // X
  CLICKS:         25, // Y
  SCORE:          26, // Z
  REPLIED:        27, // AA
  BOUNCED:        28, // AB
  SENDER_USED:    29, // AC
};

const SEQUENCE_DAYS = [0]; // Sirf pehla email, no follow-ups

// ─── RUN ONCE: Add tracking headers to your existing sheet ──
function setupTrackingColumns() {
  const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheets()[0];

  const newHeaders = [
    "Status","Seq Day","Last Sent",
    "Opens","Clicks","Score",
    "Replied","Bounced","Sender Used"
  ];

  newHeaders.forEach((header, i) => {
    const cell = sheet.getRange(1, 21 + i);
    cell.setValue(header);
    cell.setFontWeight("bold");
    cell.setBackground("#1a1a2e");
    cell.setFontColor("#ffffff");
  });

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, COL.STATUS,  lastRow - 1, 1).setValue("New");
    sheet.getRange(2, COL.SEQ_DAY, lastRow - 1, 1).setValue(0);
    sheet.getRange(2, COL.OPENS,   lastRow - 1, 1).setValue(0);
    sheet.getRange(2, COL.CLICKS,  lastRow - 1, 1).setValue(0);
    sheet.getRange(2, COL.SCORE,   lastRow - 1, 1).setValue(0);
    sheet.getRange(2, COL.REPLIED, lastRow - 1, 1).setValue(false);
    sheet.getRange(2, COL.BOUNCED, lastRow - 1, 1).setValue(false);
  }

  Logger.log("Done! Tracking columns added. Rows initialized: " + (lastRow - 1));
}

// ─── MAIN FUNCTION — Set as daily 9am trigger ────────────────
function runOutboundSystem() {
  const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheets()[0];
  const data  = sheet.getDataRange().getValues();

  const dailySentMap = {};
  CONFIG.SENDER_ACCOUNTS.forEach(acc => dailySentMap[acc.email] = 0);

  detectReplies(sheet, data);

  let totalSent = 0;

  for (let i = 1; i < data.length; i++) {
    const row    = data[i];
    const rowNum = i + 1;

    const fullName    = row[COL.FULL_NAME    - 1] || "";
    const username    = row[COL.USERNAME     - 1] || "";
    const publicEmail = row[COL.PUBLIC_EMAIL - 1] || "";
    const biography   = row[COL.BIOGRAPHY   - 1] || "";
    const category    = row[COL.CATEGORY    - 1] || "";
    const followers   = row[COL.FOLLOWERS   - 1] || 0;
    const city        = row[COL.CITY        - 1] || "";
    const instagramId = row[COL.INSTAGRAM_ID- 1] || "";

    const status  = row[COL.STATUS  - 1] || "New";
    const replied = row[COL.REPLIED - 1];
    const bounced = row[COL.BOUNCED - 1];
    const seqDay  = parseInt(row[COL.SEQ_DAY   - 1]) || 0;
    const lastSent= row[COL.LAST_SENT - 1];

    // Skip conditions
    if (!publicEmail || publicEmail.trim() === "") continue;
    if (replied || bounced) continue;
    if (status === "Completed" || status === "Replied") continue;
    if (seqDay >= SEQUENCE_DAYS.length) {
      sheet.getRange(rowNum, COL.STATUS).setValue("Completed");
      continue;
    }
    if (!shouldSendToday(lastSent, seqDay)) continue;

    const sender = pickSender(dailySentMap);
    if (!sender) { Logger.log("Daily limit reached."); break; }

    const lead = { name: fullName || username, niche: category, bio: biography, followers, city, id: instagramId };
    const emailContent = generateEmail(seqDay, lead, sender);
    if (!emailContent) continue;

    const sent = sendEmail(sender, publicEmail, emailContent);
    if (!sent) continue;

    const nextSeqDay = seqDay + 1;
    sheet.getRange(rowNum, COL.STATUS     ).setValue(nextSeqDay >= SEQUENCE_DAYS.length ? "Completed" : "Active");
    sheet.getRange(rowNum, COL.SEQ_DAY    ).setValue(nextSeqDay);
    sheet.getRange(rowNum, COL.LAST_SENT  ).setValue(new Date());
    sheet.getRange(rowNum, COL.SENDER_USED).setValue(sender.email);

    dailySentMap[sender.email]++;
    totalSent++;

    Logger.log("Sent [" + totalSent + "] to " + (fullName || username) + " <" + publicEmail + "> | Step " + (seqDay+1) + "/5");

    Utilities.sleep(Math.floor(Math.random() * (CONFIG.MAX_DELAY_MS - CONFIG.MIN_DELAY_MS) + CONFIG.MIN_DELAY_MS));
  }

  Logger.log("Session done. Total sent: " + totalSent);
}

// ─── TIMING ──────────────────────────────────────────────────
function shouldSendToday(lastSent, seqDay) {
  if (!lastSent && seqDay === 0) return true;
  if (!lastSent) return false;
  const diffDays = Math.floor((new Date() - new Date(lastSent)) / 86400000);
  const gaps = [0, 3, 2, 3, 4];
  return diffDays >= (gaps[seqDay] || 3);
}

function pickSender(dailySentMap) {
  const shuffled = [...CONFIG.SENDER_ACCOUNTS].sort(() => Math.random() - 0.5);
  for (const acc of shuffled) {
    if (dailySentMap[acc.email] < CONFIG.DAILY_LIMIT_PER_ACCOUNT) return acc;
  }
  return null;
}

// ─── EMAIL GENERATOR ─────────────────────────────────────────
function generateEmail(seqDay, lead, sender) {
  const name      = lead.name || "there";
  const niche     = (lead.niche && lead.niche.trim() !== "") ? lead.niche : "Instagram";
  const followers = lead.followers ? Number(lead.followers).toLocaleString() + " followers" : "";
  const personalized = getAIPersonalization(lead);
  const pixel = buildTrackingPixel(lead.id);

  const subjectPools = [
    ["Quick thought on your " + niche + " content", "Noticed something on your profile", niche + " growth idea"],
    ["Re: your profile", "Bumping this up", "Did this land?"],
    ["What we did for a similar " + niche + " account"],
    ["One specific thing I noticed, " + name],
    ["Closing your file"],
  ];

  const subject = subjectPools[seqDay][Math.floor(Math.random() * subjectPools[seqDay].length)];

  const bodies = [
    "Hi " + name + ",\n\n" + personalized + "\n\nMost " + niche + " accounts with " + followers + " are leaving serious inquiry volume on the table — not because of bad content, but because the profile-to-lead funnel isn't set up right.\n\nI put together a quick audit for your account. Worth a 10-minute call to walk through it?\n\n" + buildSignature(sender) + "\n" + pixel,

    "Hey " + name + ",\n\nJust bumping this up — wanted to make sure it didn't get buried.\n\nThe audit I prepared has some specific observations about your " + niche + " content funnel. Happy to share it on a quick call.\n\nWorth 10 minutes this week?\n\n" + buildSignature(sender) + "\n" + pixel,

    "Hey " + name + ",\n\nRecently worked with a " + niche + " account similar to yours. Within 60 days:\n\n→ Profile visit-to-inquiry rate: 1.4% → 7.2%\n→ 3x more DM leads, same posting frequency\n→ Zero extra ad spend\n\nThe fix was a specific funnel adjustment most accounts miss.\n\nI can walk you through exactly what we did. Tuesday or Thursday work?\n\n" + buildSignature(sender) + "\n" + pixel,

    "Hey " + name + ",\n\nLooking at your profile again — your content quality is clearly strong. But there's a specific gap between your engagement and actual inquiry volume.\n\nMost " + niche + " accounts experience this and don't realize until they see the data.\n\nWant me to send the breakdown?\n\n" + buildSignature(sender) + "\n" + pixel,

    "Hey " + name + ",\n\nI'll assume the timing isn't right — completely understand.\n\nClosing out your audit on our end. If things change and you want to explore what's possible for your brand, you know where to find me.\n\nWishing you all the best,\n\n" + buildSignature(sender),
  ];

  return { subject, body: buildHtml(bodies[seqDay]) };
}

// ─── AI PERSONALIZATION ──────────────────────────────────────
function getAIPersonalization(lead) {
  const { name, niche, bio, followers, city } = lead;
  if (!bio || bio.trim().length < 20) {
    return "Your " + niche + " brand stood out while I was reviewing accounts in this space.";
  }
  try {
    const prompt = "Write ONE personalized opening observation (max 2 sentences) for a cold email to " + name + ", a " + niche + " brand.\n\nBio: \"" + bio + "\"\nFollowers: " + followers + "\nCity: " + city + "\n\nRules:\n- Sound human, not salesy\n- Reference something specific from bio\n- Lead into: their content isn't fully converting visitors into leads\n- Do NOT start with I, Hope, or Hi\n- No greeting — just the observation";

    const res = UrlFetchApp.fetch("https://api.openai.com/v1/chat/completions", {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + CONFIG.OPENAI_API_KEY },
      payload: JSON.stringify({ model: "gpt-4o-mini", max_tokens: 80, temperature: 0.8, messages: [{ role: "user", content: prompt }] }),
      muteHttpExceptions: true,
    });
    const result = JSON.parse(res.getContentText());
    return result.choices?.[0]?.message?.content?.trim() || "";
  } catch(e) {
    return "Your " + niche + " content caught my attention while reviewing accounts in this space.";
  }
}

// ─── REPLY DETECTION ─────────────────────────────────────────
function detectReplies(sheet, data) {
  try {
    const threads = GmailApp.search("label:inbox newer_than:3d");
    const repliedEmails = new Set();
    threads.forEach(t => t.getMessages().forEach(msg => {
      const from = msg.getFrom().toLowerCase();
      const isOurs = CONFIG.SENDER_ACCOUNTS.some(a => from.includes(a.email.toLowerCase()));
      if (!isOurs) repliedEmails.add(from.replace(/.*<(.+)>/, "$1").trim());
    }));

    for (let i = 1; i < data.length; i++) {
      const email = (data[i][COL.PUBLIC_EMAIL - 1] || "").toLowerCase().trim();
      if (email && repliedEmails.has(email)) {
        sheet.getRange(i+1, COL.REPLIED).setValue(true);
        sheet.getRange(i+1, COL.STATUS ).setValue("Replied");
        updateScore(sheet, i+1, data[i], 50);
        Logger.log("Reply: " + email);
      }
    }
  } catch(e) { Logger.log("Reply detect error: " + e.message); }
}

// ─── SCORE + HELPERS ─────────────────────────────────────────
function updateScore(sheet, rowNum, row, delta) {
  const cur = parseInt(row[COL.SCORE - 1]) || 0;
  sheet.getRange(rowNum, COL.SCORE).setValue(cur + delta);
}

function buildTrackingPixel(id) {
  return '<img src="' + CONFIG.TRACKING_SERVER + '/track/open?id=' + id + '&t=' + Date.now() + '" width="1" height="1" style="display:none;"/>';
}

function buildSignature(sender) {
  return "<br><br>--<br>" + sender.alias + "<br>Nexotize Media";
}

function buildHtml(text) {
  return '<div style="font-family:Arial,sans-serif;font-size:15px;line-height:1.8;color:#1a1a1a;max-width:540px;">' + text.replace(/\n/g, "<br>") + '</div>';
}

function sendEmail(sender, toEmail, content) {
  try {
    GmailApp.sendEmail(toEmail, content.subject, "", {
      from: sender.email, name: sender.alias, htmlBody: content.body,
    });
    return true;
  } catch(e) {
    Logger.log("Send error " + toEmail + ": " + e.message);
    return false;
  }
}

// ─── SYNC TRACKING FROM SERVER ───────────────────────────────
function syncTrackingData() {
  try {
    const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheets()[0];
    const data  = sheet.getDataRange().getValues();
    const res   = UrlFetchApp.fetch(CONFIG.TRACKING_SERVER + "/api/events", { muteHttpExceptions: true });
    const events = JSON.parse(res.getContentText());

    events.forEach(event => {
      for (let i = 1; i < data.length; i++) {
        if (String(data[i][COL.INSTAGRAM_ID - 1]) !== String(event.id)) continue;
        const rowNum = i + 1;
        if (event.type === "open") {
          const opens = parseInt(data[i][COL.OPENS - 1]) || 0;
          sheet.getRange(rowNum, COL.OPENS).setValue(opens + 1);
          updateScore(sheet, rowNum, data[i], opens === 0 ? 10 : 20);
        }
        if (event.type === "click") {
          const clicks = parseInt(data[i][COL.CLICKS - 1]) || 0;
          sheet.getRange(rowNum, COL.CLICKS).setValue(clicks + 1);
          updateScore(sheet, rowNum, data[i], 30);
        }
        break;
      }
    });
    Logger.log("Tracking sync done. Events: " + events.length);
  } catch(e) { Logger.log("Sync error: " + e.message); }
}

// ─── TEST ONE EMAIL (dry run) ─────────────────────────────────
function testSendOne() {
  const sheet = SpreadsheetApp.openById(CONFIG.SHEET_ID).getSheets()[0];
  const data  = sheet.getDataRange().getValues();
  const row   = data[1]; // Row 2

  const lead = {
    name:      row[COL.FULL_NAME    - 1],
    niche:     row[COL.CATEGORY     - 1],
    bio:       row[COL.BIOGRAPHY    - 1],
    followers: row[COL.FOLLOWERS    - 1],
    city:      row[COL.CITY         - 1],
    id:        row[COL.INSTAGRAM_ID - 1],
  };

  Logger.log("Lead: " + JSON.stringify(lead));
  Logger.log("Email: " + row[COL.PUBLIC_EMAIL - 1]);

  const content = generateEmail(0, lead, CONFIG.SENDER_ACCOUNTS[0]);
  Logger.log("Subject: " + content.subject);
  Logger.log("Body:\n" + content.body.replace(/<[^>]+>/g, ""));

  // Remove comment below to actually send:
  // sendEmail(CONFIG.SENDER_ACCOUNTS[0], row[COL.PUBLIC_EMAIL - 1], content);
}
