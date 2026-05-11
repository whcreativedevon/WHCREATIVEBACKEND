// ============================================================
// WH Creative — Booking Backend
// Emails: Resend (resend.com)
// Payments: Stripe
// Calendar: Cal.com (availability) + iCal sync
// Deploy to: Railway (railway.app)
// ============================================================

const express = require('express');
const cors    = require('cors');
const Stripe  = require('stripe');
const { Resend } = require('resend');

const app    = express();
const stripe = Stripe(process.env.STRIPE_SECRET_KEY);
const resend = new Resend(process.env.RESEND_API_KEY);

app.use(cors());
app.use(express.json());

// ============================================================
// ROUTE 1 — POST /submit-booking
// Called when agent submits the booking form
// Creates Stripe payment link + sends all emails
// ============================================================
app.post('/submit-booking', async (req, res) => {
  try {
    const {
      service, serviceLabel, duration,
      date, time,
      seller, notes,
      payMethod,
      agent,
      total,
      stagingGuideUrl,
    } = req.body;

    // ── Build appointment times ──────────────────────────
    const durationHours = service === 'video' ? 3 : 1;
    const startDT       = new Date(`${date}T${time}:00`);
    const endDT         = new Date(startDT.getTime() + durationHours * 60 * 60 * 1000);

    const fmtDate  = startDT.toLocaleDateString('en-GB', { weekday:'long', day:'numeric', month:'long', year:'numeric', timeZone:'Europe/London' });
    const fmtStart = startDT.toLocaleTimeString('en-GB', { hour:'2-digit', minute:'2-digit', timeZone:'Europe/London' });
    const fmtEnd   = endDT.toLocaleTimeString('en-GB',   { hour:'2-digit', minute:'2-digit', timeZone:'Europe/London' });

    // ── Create Stripe payment link (if stripe method) ────
    let stripePaymentUrl = null;

    if (payMethod === 'stripe') {
      const session = await stripe.checkout.sessions.create({
        payment_method_types: ['card'],
        mode:           'payment',
        customer_email: seller.email,
        line_items: [{
          price_data: {
            currency:     'gbp',
            product_data: {
              name:        `WH Creative — ${serviceLabel}`,
              description: `${seller.address} | ${fmtDate} at ${fmtStart}`,
            },
            unit_amount: total * 100,
          },
          quantity: 1,
        }],
        success_url: `${process.env.BACKEND_URL}/payment-success?session_id={CHECKOUT_SESSION_ID}`,
        cancel_url:  `${process.env.BACKEND_URL}/payment-cancelled`,
        metadata: {
          sellerEmail:     seller.email,
          sellerName:      seller.name,
          agentEmail:      agent.email,
          agentName:       agent.name,
          serviceLabel,
          fmtDate,
          fmtStart,
          fmtEnd,
          address:         seller.address,
          stagingGuideUrl: stagingGuideUrl || '',
          total:           String(total),
        },
      });

      stripePaymentUrl = session.url;
    }

    // ── Send email to seller ─────────────────────────────
    await resend.emails.send({
      from:    `WH Creative <bookings@${process.env.EMAIL_DOMAIN}>`,
      to:      seller.email,
      subject: `Your property shoot — ${fmtDate}`,
      html:    sellerEmail({ seller, serviceLabel, fmtDate, fmtStart, fmtEnd, total, notes, payMethod, stripePaymentUrl, stagingGuideUrl }),
    });

    // ── Send confirmation to agent ───────────────────────
    await resend.emails.send({
      from:    `WH Creative <bookings@${process.env.EMAIL_DOMAIN}>`,
      to:      agent.email,
      subject: `Booking submitted — ${seller.name} — ${fmtDate}`,
      html:    agentEmail({ seller, agent, serviceLabel, duration, fmtDate, fmtStart, fmtEnd, total, notes, payMethod }),
    });

    // ── Send notification to WH Creative ─────────────────
    await resend.emails.send({
      from:    `WH Creative Bookings <bookings@${process.env.EMAIL_DOMAIN}>`,
      to:      process.env.OWNER_EMAIL,
      subject: `New booking — ${seller.name} — ${fmtDate}`,
      html:    ownerEmail({ seller, agent, serviceLabel, duration, fmtDate, fmtStart, fmtEnd, total, notes, payMethod }),
    });

    res.json({
      success: true,
      stripePaymentUrl,
      message: payMethod === 'stripe'
        ? 'Booking created. Payment link sent to seller.'
        : 'Booking confirmed. In-person payment on the day.',
    });

  } catch (err) {
    console.error('Booking error:', err.message);
    res.status(500).json({ error: 'Booking failed', detail: err.message });
  }
});

// ============================================================
// ROUTE 2 — GET /payment-success
// Stripe redirects here after seller pays
// Sends receipt + staging guide to seller
// ============================================================
app.get('/payment-success', async (req, res) => {
  try {
    const { session_id } = req.query;
    const session = await stripe.checkout.sessions.retrieve(session_id);

    if (session.payment_status !== 'paid') {
      return res.send('<h2>Payment not completed. Please contact your estate agent.</h2>');
    }

    const {
      sellerEmail, sellerName,
      serviceLabel, fmtDate, fmtStart,
      address, stagingGuideUrl, total,
    } = session.metadata;

    // Send receipt + staging guide to seller
    await resend.emails.send({
      from:    `WH Creative <bookings@${process.env.EMAIL_DOMAIN}>`,
      to:      sellerEmail,
      subject: `Payment confirmed — Your shoot is booked ✅`,
      html:    receiptEmail({
        sellerName, serviceLabel, fmtDate, fmtStart,
        total, address, stagingGuideUrl,
        paymentRef: session.payment_intent,
      }),
    });

    // Notify WH Creative that payment received
    await resend.emails.send({
      from:    `WH Creative Bookings <bookings@${process.env.EMAIL_DOMAIN}>`,
      to:      process.env.OWNER_EMAIL,
      subject: `✅ Payment received — ${sellerName} — ${fmtDate}`,
      html:    `<div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;">
        <p><strong>Payment confirmed</strong></p>
        <p><strong>£${total}</strong> received from <strong>${sellerName}</strong> (${sellerEmail}).</p>
        <p>Shoot on <strong>${fmtDate} at ${fmtStart}</strong> at ${address} is confirmed.</p>
        <p style="color:#aaa;font-size:12px;">Payment ref: ${session.payment_intent}</p>
      </div>`,
    });

    // Show confirmation page to seller
    res.send(`<!DOCTYPE html><html><head><meta charset="UTF-8">
      <title>Confirmed — WH Creative</title>
      <style>
        body { font-family: Arial, sans-serif; background: #f8f8f8; text-align: center; padding: 60px; color: #111; }
        .card { background: #fff; max-width: 500px; margin: 0 auto; padding: 48px; border-radius: 8px; box-shadow: 0 2px 20px rgba(0,0,0,.08); }
        .tick { font-size: 60px; margin-bottom: 16px; }
        h1 { color: #2F2F2F; margin-bottom: 12px; }
        p { color: #888; line-height: 1.7; }
        a { color: #FFA970; font-weight: bold; }
        .ref { font-size: 12px; color: #ccc; margin-top: 24px; }
      </style>
    </head><body>
      <div class="card">
        <div class="tick">✅</div>
        <h1>Payment confirmed!</h1>
        <p>Thank you <strong>${sellerName}</strong>.<br>
        Your shoot is confirmed for <strong>${fmtDate} at ${fmtStart}</strong>.<br><br>
        A receipt and home staging guide have been sent to ${sellerEmail}.</p>
        ${stagingGuideUrl && stagingGuideUrl !== 'YOUR_STAGING_GUIDE_PDF_URL'
          ? `<p style="margin-top:24px;"><a href="${stagingGuideUrl}" target="_blank">Download your home staging guide</a></p>` : ''}
        <p class="ref">Ref: ${session.payment_intent}</p>
      </div>
    </body></html>`);

  } catch (err) {
    console.error('Payment success error:', err.message);
    res.status(500).send('<h2>Something went wrong. Please contact info@wh-creative.co.uk</h2>');
  }
});

// ============================================================
// ROUTE 3 — POST /stripe-webhook
// Backup from Stripe in case the redirect fails
// ============================================================
app.post('/stripe-webhook', express.raw({ type: 'application/json' }), (req, res) => {
  try {
    const event = stripe.webhooks.constructEvent(
      req.body,
      req.headers['stripe-signature'],
      process.env.STRIPE_WEBHOOK_SECRET
    );
    if (event.type === 'checkout.session.completed') {
      console.log('Webhook: payment confirmed for', event.data.object.customer_email);
    }
    res.json({ received: true });
  } catch (err) {
    console.error('Webhook error:', err.message);
    res.status(400).send(`Webhook error: ${err.message}`);
  }
});

// Payment cancelled page
app.get('/payment-cancelled', (req, res) => {
  res.send(`<!DOCTYPE html><html><head><title>Cancelled — WH Creative</title>
  <style>body{font-family:Arial,sans-serif;text-align:center;padding:60px;background:#f8f8f8;}
  .card{background:#fff;max-width:480px;margin:0 auto;padding:48px;border-radius:8px;}
  h1{color:#2F2F2F;}p{color:#888;line-height:1.7;}</style>
  </head><body><div class="card">
  <h1>Payment not completed</h1>
  <p>Your booking has not been confirmed.<br>
  Please contact your estate agent to retry,<br>
  or email <a href="mailto:info@wh-creative.co.uk" style="color:#FFA970;font-weight:bold;">info@wh-creative.co.uk</a> for help.</p>
  </div></body></html>`);
});

// Health check
app.get('/', (req, res) => res.json({
  status:  'WH Creative backend running ✅',
  edition: 'Resend email edition',
}));

// ============================================================
// EMAIL TEMPLATES
// ============================================================

function sellerEmail({ seller, serviceLabel, fmtDate, fmtStart, fmtEnd, total, notes, payMethod, stripePaymentUrl, stagingGuideUrl }) {
  const payBlock = payMethod === 'stripe'
    ? `<div style="text-align:center;margin:32px 0;">
        <a href="${stripePaymentUrl}" style="background:#FFA970;color:#111;padding:16px 40px;border-radius:4px;text-decoration:none;font-weight:bold;font-size:16px;display:inline-block;">
          Pay now &mdash; &pound;${total}
        </a>
        <p style="color:#aaa;font-size:12px;margin-top:12px;">Secure payment powered by Stripe.<br>Your booking is confirmed once payment is received.</p>
      </div>`
    : `<p style="background:#f0faf4;border:1px solid #b8ddc5;border-radius:4px;padding:14px 18px;color:#2d6a45;">
        <strong>Payment:</strong> &pound;${total} will be taken in person at the property on the day of your shoot.
      </p>`;

  const stageBlock = stagingGuideUrl && stagingGuideUrl !== 'YOUR_STAGING_GUIDE_PDF_URL'
    ? `<div style="background:#fff8f3;border:1px solid #FFD4B3;border-radius:4px;padding:16px 20px;margin:24px 0;">
        <p style="margin:0;font-weight:bold;color:#e8894a;">&#127968; Home staging guide</p>
        <p style="margin:8px 0 0;font-size:14px;color:#555;">Read our guide before the day to get the best results from your shoot.</p>
        <p style="margin:10px 0 0;"><a href="${stagingGuideUrl}" style="color:#e8894a;font-weight:bold;">Download the guide &rarr;</a></p>
      </div>` : '';

  return wrap('WH Creative', 'Property Media', `
    <p>Dear ${seller.name},</p>
    <p>Your property shoot has been arranged with WH Creative. Here are your details:</p>
    ${table([
      ['Service',  serviceLabel],
      ['Date',     fmtDate],
      ['Time',     `${fmtStart} &ndash; ${fmtEnd}`],
      ['Property', seller.address],
      ['Total',    `&pound;${total}`],
    ])}
    ${notes ? note(notes) : ''}
    ${payBlock}
    ${stageBlock}
    <p style="color:#555;">Any questions, please get in touch.</p>
    <p style="color:#555;">Kind regards,<br><strong>WH Creative</strong><br>
    <a href="mailto:info@wh-creative.co.uk" style="color:#e8894a;">info@wh-creative.co.uk</a></p>
  `, true);
}

function agentEmail({ seller, agent, serviceLabel, duration, fmtDate, fmtStart, fmtEnd, total, notes, payMethod }) {
  return wrap('WH Creative', `Booking confirmation for ${agent.name}`, `
    <p>Hi ${agent.name},</p>
    <p>The following booking has been submitted successfully.</p>
    ${table([
      ['Service',   serviceLabel],
      ['Duration',  duration],
      ['Date',      fmtDate],
      ['Time',      `${fmtStart} &ndash; ${fmtEnd}`],
      ['Property',  seller.address],
      ['Seller',    `${seller.name}<br>${seller.email}<br>${seller.phone}<br>${seller.bedrooms} bedrooms`],
      ['Payment',   payMethod === 'stripe' ? '&#8987; Stripe link sent to seller &mdash; awaiting payment' : '&#128179; In-person card on the day'],
      ['Total',     `&pound;${total}`],
    ])}
    ${notes ? note(notes) : ''}
    <p style="color:#aaa;font-size:13px;">${payMethod === 'stripe' ? 'The booking is confirmed once the seller completes payment.' : 'The booking is confirmed.'}</p>
    <p>Kind regards,<br><strong>WH Creative</strong></p>
  `);
}

function ownerEmail({ seller, agent, serviceLabel, duration, fmtDate, fmtStart, fmtEnd, total, notes, payMethod }) {
  return wrap('New booking received',
    payMethod === 'stripe' ? '&#8987; Awaiting seller payment' : '&#9989; Confirmed — in-person payment', `
    ${table([
      ['Service',    `${serviceLabel} (${duration})`],
      ['Date & time', `${fmtDate}<br>${fmtStart} &ndash; ${fmtEnd}`],
      ['Property',   seller.address],
      ['Seller',     `${seller.name}<br>${seller.email}<br>${seller.phone}<br>${seller.bedrooms} bedrooms`],
      ['Booked by',  `${agent.name}<br>${agent.email}`],
      ['Payment',    payMethod === 'stripe' ? '&#8987; Awaiting Stripe payment' : '&#128179; In-person card on the day'],
      ['Total',      `&pound;${total}`],
    ])}
    ${notes ? note(notes) : ''}
  `);
}

function receiptEmail({ sellerName, serviceLabel, fmtDate, fmtStart, total, address, stagingGuideUrl, paymentRef }) {
  const stageBlock = stagingGuideUrl && stagingGuideUrl !== 'YOUR_STAGING_GUIDE_PDF_URL'
    ? `<div style="background:#fff8f3;border:1px solid #FFD4B3;border-radius:4px;padding:16px 20px;margin:24px 0;">
        <p style="margin:0;font-weight:bold;color:#e8894a;">&#127968; Your home staging guide</p>
        <p style="margin:8px 0 0;font-size:14px;color:#555;">Make the most of your shoot — read our guide before the day.</p>
        <p style="margin:10px 0 0;"><a href="${stagingGuideUrl}" style="color:#e8894a;font-weight:bold;">Download your guide &rarr;</a></p>
      </div>` : '';

  return wrap('Payment confirmed &#9989;', 'WH Creative &mdash; Property Media', `
    <p>Dear ${sellerName},</p>
    <p>Your payment has been received and your shoot is confirmed. Here is your receipt:</p>
    ${table([
      ['Service',      serviceLabel],
      ['Date',         fmtDate],
      ['Time',         fmtStart],
      ['Property',     address],
      ['Amount paid',  `<span style="color:#2d6a45;font-size:18px;font-weight:bold;">&pound;${total}</span>`],
      ['Payment ref',  `<span style="font-size:12px;color:#aaa;">${paymentRef}</span>`],
    ])}
    ${stageBlock}
    <p style="color:#555;">We look forward to seeing you on the day.<br>
    Any questions: <a href="mailto:info@wh-creative.co.uk" style="color:#e8894a;">info@wh-creative.co.uk</a></p>
    <p style="color:#555;">Kind regards,<br><strong>WH Creative</strong></p>
  `, true);
}

// ── Template helpers ──────────────────────────────────────

function wrap(title, subtitle, content, gdpr = false) {
  return `<div style="font-family:Arial,sans-serif;max-width:600px;margin:0 auto;color:#111;">
    <div style="background:#2F2F2F;padding:24px 32px;border-radius:6px 6px 0 0;">
      <p style="font-size:22px;color:#FFA970;margin:0;font-weight:bold;">${title}</p>
      <p style="color:#aaa;margin:4px 0 0;font-size:12px;">${subtitle}</p>
    </div>
    <div style="background:#fff;padding:32px;border:1px solid #e0e0e0;border-top:none;border-radius:0 0 6px 6px;">
      ${content}
    </div>
    ${gdpr ? `<p style="text-align:center;font-size:11px;color:#ccc;margin-top:16px;">
      Your data is handled in accordance with UK GDPR.
      <a href="mailto:info@wh-creative.co.uk" style="color:#ccc;">Contact us</a> to request access or deletion.
    </p>` : ''}
  </div>`;
}

function table(rows) {
  return `<table style="width:100%;border-collapse:collapse;margin:20px 0;font-size:14px;">
    ${rows.map(([l, v], i) => `
      <tr>
        <td style="padding:11px;border-bottom:${i < rows.length - 1 ? '1px solid #f0f0f0' : 'none'};color:#888;width:38%;vertical-align:top;">${l}</td>
        <td style="padding:11px;border-bottom:${i < rows.length - 1 ? '1px solid #f0f0f0' : 'none'};vertical-align:top;">${v}</td>
      </tr>`).join('')}
  </table>`;
}

function note(n) {
  return `<p style="background:#f8f8f8;padding:12px 16px;border-radius:4px;font-size:14px;"><strong>Access notes:</strong> ${n}</p>`;
}

// ── Start server ──────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`WH Creative backend running on port ${PORT}`));
