// ============================================================
// WH Creative — Booking Backend (Microsoft 365 / Outlook)
// Deploy to Railway (railway.app) — free tier
// Uses Microsoft Graph API for calendar + email
// ============================================================

const express = require('express');
const cors    = require('cors');
const Stripe  = require('stripe');
const fetch   = require('node-fetch');

const app = express();
app.use(cors());
app.use(express.json());

const stripe = Stripe(process.env.STRIPE_SECRET_KEY);

// ============================================================
// MICROSOFT GRAPH HELPERS
// ============================================================

async function getGraphToken() {
  const url  = `https://login.microsoftonline.com/${process.env.MS_TENANT_ID}/oauth2/v2.0/token`;
  const body = new URLSearchParams({
    grant_type:    'client_credentials',
    client_id:     process.env.MS_CLIENT_ID,
    client_secret: process.env.MS_CLIENT_SECRET,
    scope:         'https://graph.microsoft.com/.default',
  });
  const res  = await fetch(url, { method: 'POST', body });
  const data = await res.json();
  if (!data.access_token) throw new Error('Graph token failed: ' + JSON.stringify(data));
  return data.access_token;
}

async function graphCall(method, path, body = null) {
  const token = await getGraphToken();
  const res   = await fetch(`https://graph.microsoft.com/v1.0${path}`, {
    method,
    headers: { 'Authorization': `Bearer ${token}`, 'Content-Type': 'application/json' },
    body: body ? JSON.stringify(body) : undefined,
  });
  if (method === 'DELETE' || res.status === 204) return null;
  const data = await res.json();
  if (!res.ok) throw new Error(`Graph ${res.status}: ${JSON.stringify(data)}`);
  return data;
}

async function sendEmail(from, { to, subject, html }) {
  await graphCall('POST', `/users/${from}/sendMail`, {
    message: {
      subject,
      body: { contentType: 'HTML', content: html },
      toRecipients: [{ emailAddress: { address: to } }],
    },
    saveToSentItems: true,
  });
}

// ============================================================
// ROUTE 1 — GET /available-slots
// Reads Outlook calendar and returns booked times for a date
// ============================================================
app.get('/available-slots', async (req, res) => {
  try {
    const { date } = req.query;
    if (!date) return res.status(400).json({ error: 'date param required' });

    const user = process.env.OUTLOOK_EMAIL;
    const data = await graphCall('GET',
      `/users/${user}/calendarView?startDateTime=${date}T00:00:00&endDateTime=${date}T23:59:59&$select=subject,start,end&$orderby=start/dateTime`
    );

    const bookedSlots = (data.value || []).map(ev => {
      if (!ev.start || !ev.start.dateTime) return null;
      // Graph returns UTC — convert to UK time
      const utc    = new Date(ev.start.dateTime + (ev.start.dateTime.endsWith('Z') ? '' : 'Z'));
      const london = new Date(utc.toLocaleString('en-GB', { timeZone: 'Europe/London' }));
      return String(london.getHours()).padStart(2,'0') + ':' + String(london.getMinutes()).padStart(2,'0');
    }).filter(Boolean);

    res.json({ date, bookedSlots });
  } catch (err) {
    console.error('Slots error:', err.message);
    res.status(500).json({ error: 'Failed to fetch availability', detail: err.message });
  }
});

// ============================================================
// ROUTE 2 — POST /submit-booking
// Creates calendar event, Stripe link, and sends all emails
// ============================================================
app.post('/submit-booking', async (req, res) => {
  try {
    const { service, serviceLabel, duration, date, time, seller, notes, payMethod, agent, total, stagingGuideUrl } = req.body;
    const user          = process.env.OUTLOOK_EMAIL;
    const durationHours = service === 'video' ? 3 : 1;

    const startDT = new Date(`${date}T${time}:00`);
    const endDT   = new Date(startDT.getTime() + durationHours * 60 * 60 * 1000);

    const fmtDate  = startDT.toLocaleDateString('en-GB',  { weekday:'long', day:'numeric', month:'long', year:'numeric', timeZone:'Europe/London' });
    const fmtStart = startDT.toLocaleTimeString('en-GB',  { hour:'2-digit', minute:'2-digit', timeZone:'Europe/London' });
    const fmtEnd   = endDT.toLocaleTimeString('en-GB',    { hour:'2-digit', minute:'2-digit', timeZone:'Europe/London' });

    // 1. Create Outlook calendar event
    const calEvent = {
      subject:  `📷 ${serviceLabel} — ${seller.address}`,
      body:     { contentType: 'text', content: [
        `Service: ${serviceLabel} (${duration})`,
        `Seller: ${seller.name} | ${seller.phone} | ${seller.email}`,
        `Bedrooms: ${seller.bedrooms}`,
        `Booked by: ${agent.name} (${agent.email})`,
        `Payment: ${payMethod === 'stripe' ? 'AWAITING STRIPE PAYMENT from seller' : 'In-person card on the day'}`,
        `Total: £${total}`,
        notes ? `Notes: ${notes}` : '',
      ].filter(Boolean).join('\n') },
      location: { displayName: seller.address },
      start:    { dateTime: `${date}T${time}:00`, timeZone: 'GMT Standard Time' },
      end:      { dateTime: endDT.toISOString().slice(0,19), timeZone: 'UTC' },
      showAs:   'busy',
      isReminderOn: true,
      reminderMinutesBeforeStart: 60,
    };

    const created = await graphCall('POST', `/users/${user}/events`, calEvent);
    const eventId = created.id;

    // 2. Create Stripe payment link if needed
    let stripePaymentUrl = null;
    if (payMethod === 'stripe') {
      const session = await stripe.checkout.sessions.create({
        payment_method_types: ['card'],
        mode: 'payment',
        customer_email: seller.email,
        line_items: [{ price_data: {
          currency: 'gbp',
          product_data: { name: `WH Creative — ${serviceLabel}`, description: `${seller.address} | ${fmtDate} at ${fmtStart}` },
          unit_amount: total * 100,
        }, quantity: 1 }],
        success_url: `${process.env.BACKEND_URL}/payment-success?session_id={CHECKOUT_SESSION_ID}&event_id=${encodeURIComponent(eventId)}`,
        cancel_url:  `${process.env.BACKEND_URL}/payment-cancelled`,
        metadata: { eventId, sellerEmail: seller.email, sellerName: seller.name, agentEmail: agent.email, agentName: agent.name, serviceLabel, fmtDate, fmtStart, address: seller.address, stagingGuideUrl: stagingGuideUrl || '', total: String(total) },
      });
      stripePaymentUrl = session.url;
    }

    // 3. Send emails
    await sendEmail(user, { to: seller.email, subject: `Your property shoot — ${fmtDate}`, html: sellerEmail({ seller, serviceLabel, fmtDate, fmtStart, fmtEnd, total, notes, payMethod, stripePaymentUrl, stagingGuideUrl }) });
    await sendEmail(user, { to: agent.email,  subject: `Booking submitted — ${seller.name} — ${fmtDate}`, html: agentEmail({ seller, agent, serviceLabel, duration, fmtDate, fmtStart, fmtEnd, total, notes, payMethod }) });
    await sendEmail(user, { to: user,          subject: `New booking — ${seller.name} — ${fmtDate}`, html: ownerEmail({ seller, agent, serviceLabel, duration, fmtDate, fmtStart, fmtEnd, total, notes, payMethod }) });

    res.json({ success: true, eventId, stripePaymentUrl });
  } catch (err) {
    console.error('Booking error:', err.message);
    res.status(500).json({ error: 'Booking failed', detail: err.message });
  }
});

// ============================================================
// ROUTE 3 — GET /payment-success
// Stripe redirects here after seller pays
// Updates calendar + sends receipt + staging guide
// ============================================================
app.get('/payment-success', async (req, res) => {
  try {
    const { session_id, event_id } = req.query;
    const user    = process.env.OUTLOOK_EMAIL;
    const session = await stripe.checkout.sessions.retrieve(session_id);

    if (session.payment_status !== 'paid') {
      return res.send('<h2>Payment not completed. Please contact your estate agent.</h2>');
    }

    const { sellerEmail, sellerName, serviceLabel, fmtDate, fmtStart, address, stagingGuideUrl, total } = session.metadata;

    // Update calendar event to confirmed
    try {
      await graphCall('PATCH', `/users/${user}/events/${event_id}`, {
        subject: `✅ ${serviceLabel} — ${address}`,
        body:    { contentType: 'text', content: `PAYMENT CONFIRMED £${total}\nSeller: ${sellerName} (${sellerEmail})\nDate: ${fmtDate} at ${fmtStart}\nRef: ${session.payment_intent}` },
      });
    } catch (e) { console.error('Calendar update error:', e.message); }

    // Send receipt to seller
    await sendEmail(user, { to: sellerEmail, subject: `Payment confirmed — Your shoot is booked ✅`, html: receiptEmail({ sellerName, serviceLabel, fmtDate, fmtStart, total, address, stagingGuideUrl, paymentRef: session.payment_intent }) });

    // Notify owner
    await sendEmail(user, { to: user, subject: `✅ Payment received — ${sellerName} — ${fmtDate}`, html: `<p><strong>£${total}</strong> received from <strong>${sellerName}</strong> (${sellerEmail}).<br>Booking on <strong>${fmtDate} at ${fmtStart}</strong> at ${address} is confirmed.<br>Ref: ${session.payment_intent}</p>` });

    res.send(`<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Confirmed — WH Creative</title>
    <style>body{font-family:Arial,sans-serif;background:#f8f8f8;text-align:center;padding:60px;}
    .card{background:#fff;max-width:500px;margin:0 auto;padding:48px;border-radius:8px;box-shadow:0 2px 20px rgba(0,0,0,.08);}
    .tick{font-size:60px;margin-bottom:16px;}h1{color:#2F2F2F;}p{color:#888;line-height:1.7;}a{color:#FFA970;}</style>
    </head><body><div class="card">
    <div class="tick">✅</div><h1>Payment confirmed!</h1>
    <p>Thank you <strong>${sellerName}</strong>.<br>Your shoot is confirmed for <strong>${fmtDate} at ${fmtStart}</strong>.<br>A receipt and home staging guide have been sent to ${sellerEmail}.</p>
    ${stagingGuideUrl && stagingGuideUrl !== 'YOUR_STAGING_GUIDE_PDF_URL' ? `<p><a href="${stagingGuideUrl}" target="_blank">Download your home staging guide</a></p>` : ''}
    </div></body></html>`);
  } catch (err) {
    console.error('Payment success error:', err.message);
    res.status(500).send('<h2>Something went wrong. Please contact info@wh-creative.co.uk</h2>');
  }
});

// ============================================================
// ROUTE 4 — Stripe webhook (backup)
// ============================================================
app.post('/stripe-webhook', express.raw({ type: 'application/json' }), (req, res) => {
  try {
    const event = stripe.webhooks.constructEvent(req.body, req.headers['stripe-signature'], process.env.STRIPE_WEBHOOK_SECRET);
    if (event.type === 'checkout.session.completed') console.log('Payment confirmed via webhook:', event.data.object.customer_email);
    res.json({ received: true });
  } catch (err) {
    res.status(400).send(`Webhook error: ${err.message}`);
  }
});

app.get('/payment-cancelled', (req, res) => res.send(`<!DOCTYPE html><html><head><title>Cancelled — WH Creative</title>
  <style>body{font-family:Arial,sans-serif;text-align:center;padding:60px;background:#f8f8f8;}
  .card{background:#fff;max-width:480px;margin:0 auto;padding:48px;border-radius:8px;}h1{color:#2F2F2F;}p{color:#888;}</style>
  </head><body><div class="card"><h1>Payment not completed</h1>
  <p>Your booking has not been confirmed. Please contact your estate agent to retry,<br>or email <a href="mailto:info@wh-creative.co.uk" style="color:#FFA970;">info@wh-creative.co.uk</a></p>
  </div></body></html>`));

app.get('/', (req, res) => res.json({ status: 'WH Creative backend running ✅', edition: 'Microsoft Outlook' }));

// ============================================================
// EMAIL TEMPLATES
// ============================================================

function sellerEmail({ seller, serviceLabel, fmtDate, fmtStart, fmtEnd, total, notes, payMethod, stripePaymentUrl, stagingGuideUrl }) {
  const payBlock = payMethod === 'stripe'
    ? `<div style="text-align:center;margin:32px 0;"><a href="${stripePaymentUrl}" style="background:#FFA970;color:#111;padding:16px 36px;border-radius:4px;text-decoration:none;font-weight:bold;font-size:16px;display:inline-block;">Pay now &mdash; &pound;${total}</a><p style="color:#aaa;font-size:12px;margin-top:10px;">Secure payment via Stripe. Booking confirmed once payment is received.</p></div>`
    : `<p style="background:#f0faf4;border:1px solid #b8ddc5;border-radius:4px;padding:14px 18px;color:#2d6a45;"><strong>Payment:</strong> &pound;${total} taken in person at the property on the day.</p>`;
  const stageBlock = stagingGuideUrl && stagingGuideUrl !== 'YOUR_STAGING_GUIDE_PDF_URL'
    ? `<div style="background:#fff8f3;border:1px solid #FFD4B3;border-radius:4px;padding:16px 20px;margin:24px 0;"><p style="margin:0;font-weight:bold;color:#e8894a;">&#127968; Home staging guide</p><p style="margin:8px 0 0;font-size:14px;color:#555;">Read our guide before the day to get the best results from your shoot.</p><p style="margin:10px 0 0;"><a href="${stagingGuideUrl}" style="color:#e8894a;font-weight:bold;">Download the guide &rarr;</a></p></div>` : '';
  return wrap('WH Creative', 'Property Media', `<p>Dear ${seller.name},</p><p>Your property shoot has been arranged. Here are your details:</p>${table([['Service',serviceLabel],['Date',fmtDate],['Time',`${fmtStart} &ndash; ${fmtEnd}`],['Property',seller.address],['Total',`&pound;${total}`]])}${notes?note(notes):''}${payBlock}${stageBlock}<p>Any questions, get in touch.</p><p>Kind regards,<br><strong>WH Creative</strong><br><a href="mailto:info@wh-creative.co.uk" style="color:#e8894a;">info@wh-creative.co.uk</a></p>`, true);
}

function agentEmail({ seller, agent, serviceLabel, duration, fmtDate, fmtStart, fmtEnd, total, notes, payMethod }) {
  return wrap('WH Creative', `Booking confirmation for ${agent.name}`, `<p>Hi ${agent.name},</p><p>The following booking has been submitted.</p>${table([['Service',serviceLabel],['Duration',duration],['Date',fmtDate],['Time',`${fmtStart} &ndash; ${fmtEnd}`],['Property',seller.address],['Seller',`${seller.name}<br>${seller.email}<br>${seller.phone}<br>${seller.bedrooms} bedrooms`],['Payment',payMethod==='stripe'?'&#8987; Stripe link sent to seller &mdash; awaiting payment':'&#128179; In-person card on the day'],['Total',`&pound;${total}`]])}${notes?note(notes):''}<p style="color:#aaa;font-size:13px;">${payMethod==='stripe'?'Booking confirmed in diary once seller pays.':'Booking confirmed in diary.'}</p><p>Kind regards,<br><strong>WH Creative</strong></p>`);
}

function ownerEmail({ seller, agent, serviceLabel, duration, fmtDate, fmtStart, fmtEnd, total, notes, payMethod }) {
  return wrap('New booking received', payMethod==='stripe'?'&#8987; Awaiting seller payment':'&#9989; Confirmed', `${table([['Service',`${serviceLabel} (${duration})`],['Date &amp; time',`${fmtDate}<br>${fmtStart} &ndash; ${fmtEnd}`],['Property',seller.address],['Seller',`${seller.name}<br>${seller.email}<br>${seller.phone}<br>${seller.bedrooms} bedrooms`],['Booked by',`${agent.name}<br>${agent.email}`],['Payment',payMethod==='stripe'?'&#8987; Awaiting Stripe payment':'&#128179; In-person card'],['Total',`&pound;${total}`]])}${notes?note(notes):''}`);
}

function receiptEmail({ sellerName, serviceLabel, fmtDate, fmtStart, total, address, stagingGuideUrl, paymentRef }) {
  const stageBlock = stagingGuideUrl && stagingGuideUrl !== 'YOUR_STAGING_GUIDE_PDF_URL'
    ? `<div style="background:#fff8f3;border:1px solid #FFD4B3;border-radius:4px;padding:16px 20px;margin:24px 0;"><p style="margin:0;font-weight:bold;color:#e8894a;">&#127968; Your home staging guide</p><p style="margin:8px 0 0;font-size:14px;color:#555;">Make the most of your shoot — read our guide before the day.</p><p style="margin:10px 0 0;"><a href="${stagingGuideUrl}" style="color:#e8894a;font-weight:bold;">Download your guide &rarr;</a></p></div>` : '';
  return wrap('Payment confirmed &#9989;', 'WH Creative &mdash; Property Media', `<p>Dear ${sellerName},</p><p>Your payment has been received and your shoot is confirmed. Here is your receipt:</p>${table([['Service',serviceLabel],['Date',fmtDate],['Time',fmtStart],['Property',address],['Amount paid',`<span style="color:#2d6a45;font-size:18px;font-weight:bold;">&pound;${total}</span>`],['Payment ref',`<span style="font-size:12px;color:#aaa;">${paymentRef}</span>`]])}${stageBlock}<p>We look forward to seeing you on the day.<br>Any questions: <a href="mailto:info@wh-creative.co.uk" style="color:#e8894a;">info@wh-creative.co.uk</a></p><p>Kind regards,<br><strong>WH Creative</strong></p>`, true);
}

// Template helpers
function wrap(title, subtitle, content, gdpr = false) {
  return `<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;color:#111;">
    <div style="background:#2F2F2F;padding:24px 32px;border-radius:6px 6px 0 0;">
      <p style="font-size:22px;color:#FFA970;margin:0;font-weight:bold;">${title}</p>
      <p style="color:#aaa;margin:4px 0 0;font-size:12px;">${subtitle}</p>
    </div>
    <div style="background:#fff;padding:32px;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 6px 6px;">${content}</div>
    ${gdpr ? '<p style="text-align:center;font-size:11px;color:#ccc;margin-top:16px;">Your data is handled in accordance with UK GDPR. <a href="mailto:info@wh-creative.co.uk" style="color:#ccc;">Contact us</a> to request access or deletion.</p>' : ''}
  </div>`;
}

function table(rows) {
  return `<table style="width:100%;border-collapse:collapse;margin:20px 0;font-size:14px;">${rows.map(([l,v],i)=>`<tr><td style="padding:11px;border-bottom:${i<rows.length-1?'1px solid #f0f0f0':'none'};color:#888;width:38%;vertical-align:top;">${l}</td><td style="padding:11px;border-bottom:${i<rows.length-1?'1px solid #f0f0f0':'none'};vertical-align:top;">${v}</td></tr>`).join('')}</table>`;
}

function note(n) {
  return `<p style="background:#f8f8f8;padding:12px 16px;border-radius:4px;font-size:14px;"><strong>Access notes:</strong> ${n}</p>`;
}

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`WH Creative backend running on port ${PORT}`));
