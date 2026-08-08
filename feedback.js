// Cloudflare Pages Function — lives at /api/feedback
//
// Set these as Cloudflare Pages environment variables (Settings > Environment
// variables). Mark MS_CLIENT_SECRET as a "Secret," not plain text.
//
//   MS_TENANT_ID      Directory (tenant) ID from the app registration
//   MS_CLIENT_ID      Application (client) ID from the app registration
//   MS_CLIENT_SECRET  The client secret VALUE (not the secret ID)
//   MS_SENDER_MAILBOX The mailbox the app sends AS — hello@fivem.group
//
// The recipient address is read from the submitted payload (managerEmail),
// which the front-end already sets per location. Right now that's the same
// hello@fivem.group for every location, but the function doesn't assume that,
// so nothing here needs to change if that ever splits out per restaurant.

export async function onRequestPost(context) {
  const { request, env } = context;

  let payload;
  try {
    payload = await request.json();
  } catch (err) {
    return json({ error: 'Invalid JSON body' }, 400);
  }

  const { location, managerEmail, rating, comment, wantsCallback, contact, timestamp } = payload;

  if (!managerEmail) {
    return json({ error: 'Missing managerEmail in payload' }, 400);
  }

  // 1. Exchange the client credentials for a Graph access token.
  let accessToken;
  try {
    const tokenRes = await fetch(
      `https://login.microsoftonline.com/${env.MS_TENANT_ID}/oauth2/v2.0/token`,
      {
        method: 'POST',
        headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
        body: new URLSearchParams({
          client_id: env.MS_CLIENT_ID,
          client_secret: env.MS_CLIENT_SECRET,
          scope: 'https://graph.microsoft.com/.default',
          grant_type: 'client_credentials'
        })
      }
    );

    if (!tokenRes.ok) {
      const detail = await tokenRes.text();
      return json({ error: 'Token request failed', detail }, 502);
    }

    const tokenData = await tokenRes.json();
    accessToken = tokenData.access_token;
  } catch (err) {
    return json({ error: 'Token request threw', detail: String(err) }, 502);
  }

  // 2. Build the email body.
  const subject = `Feedback: ${location || 'unknown location'} — ${rating || '?'}★${wantsCallback ? ' (callback requested)' : ''}`;

  const bodyLines = [
    `Location: ${location || 'unknown'}`,
    `Rating: ${rating || 'n/a'} stars`,
    `Comment: ${comment ? comment : '(none provided)'}`,
    `Wants callback: ${wantsCallback ? 'Yes' : 'No'}`,
    wantsCallback ? `Contact: ${contact || '(not provided)'}` : null,
    `Submitted: ${timestamp || new Date().toISOString()}`
  ].filter(Boolean).join('\n');

  const mailPayload = {
    message: {
      subject,
      body: { contentType: 'Text', content: bodyLines },
      toRecipients: [{ emailAddress: { address: managerEmail } }]
    },
    saveToSentItems: 'false'
  };

  // 3. Send it via Graph, as the configured sender mailbox.
  try {
    const sendRes = await fetch(
      `https://graph.microsoft.com/v1.0/users/${env.MS_SENDER_MAILBOX}/sendMail`,
      {
        method: 'POST',
        headers: {
          Authorization: `Bearer ${accessToken}`,
          'Content-Type': 'application/json'
        },
        body: JSON.stringify(mailPayload)
      }
    );

    if (!sendRes.ok) {
      const detail = await sendRes.text();
      return json({ error: 'Send failed', detail }, 502);
    }
  } catch (err) {
    return json({ error: 'Send threw', detail: String(err) }, 502);
  }

  return json({ success: true }, 200);
}

function json(body, status) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'Content-Type': 'application/json' }
  });
}
