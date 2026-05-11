# WH Creative — Backend Setup Guide
## Microsoft Outlook Edition

No Google account needed. Everything runs through your Microsoft 365 / Outlook account.

---

## What this backend does

- **Real-time availability** — reads your Outlook calendar before showing time slots to agents
- **Creates calendar events** — new bookings appear in your Outlook diary instantly
- **Stripe payment links** — sent directly to the seller's email
- **Confirms on payment** — updates the calendar event and sends receipt + staging guide
- **All emails** — sent from `info@wh-creative.co.uk` via Outlook

---

## What you need to set up

| Step | Task | Time |
|------|------|------|
| 1 | Deploy to Railway | 10 mins |
| 2 | Azure app registration | 10 mins |
| 3 | Add environment variables | 5 mins |
| 4 | Stripe webhook | 5 mins |
| 5 | Update booking form | 5 mins |

---

## STEP 1 — Deploy to Railway

1. Go to **[github.com](https://github.com)** and create a free account if you don't have one
2. Click **New repository** → name it `wh-creative-backend` → Create
3. Upload these three files to it: `server.js`, `package.json`, `.env.example`
4. Go to **[railway.app](https://railway.app)** → sign up with your GitHub account
5. Click **New Project → Deploy from GitHub repo** → select `wh-creative-backend`
6. Railway deploys it automatically (takes about 2 minutes)
7. Go to your project → **Settings → Networking → Generate Domain**
8. Copy your Railway URL — it looks like `https://wh-creative-backend-production.up.railway.app`
   **Save this — you need it in steps below**

---

## STEP 2 — Azure App Registration (Microsoft 365 access)

This gives your backend permission to read your Outlook calendar and send emails.

1. Go to **[portal.azure.com](https://portal.azure.com)** and sign in with your Microsoft 365 account (`info@wh-creative.co.uk`)
2. In the search bar type **App registrations** → click it → **New registration**
3. Fill in:
   - Name: `WH Creative Booking`
   - Supported account types: **Accounts in this organizational directory only**
   - Click **Register**
4. On the overview page, copy these two values — you'll need them:
   - **Application (client) ID** → this is your `MS_CLIENT_ID`
   - **Directory (tenant) ID** → this is your `MS_TENANT_ID`

5. Go to **Certificates & secrets** → **New client secret**
   - Description: `WH Creative Backend`
   - Expires: **24 months**
   - Click **Add**
   - Copy the **Value** immediately — this is your `MS_CLIENT_SECRET`
     *(It disappears after you leave the page)*

6. Go to **API permissions** → **Add a permission** → **Microsoft Graph** → **Application permissions**
7. Search for and add these three permissions:
   - `Calendars.ReadWrite`
   - `Mail.Send`
   - `User.Read.All`
8. Click **Grant admin consent for [your organisation]** → confirm Yes
   *(You need to be an admin on your Microsoft 365 account to do this. If you're a solo user on Microsoft 365, you are the admin.)*

---

## STEP 3 — Add environment variables to Railway

In Railway, go to your project → **Variables** → add each one:

| Variable | Value |
|----------|-------|
| `MS_TENANT_ID` | From Step 2 |
| `MS_CLIENT_ID` | From Step 2 |
| `MS_CLIENT_SECRET` | From Step 2 |
| `OUTLOOK_EMAIL` | `info@wh-creative.co.uk` |
| `STRIPE_SECRET_KEY` | Your Stripe secret key (`sk_live_...`) |
| `STRIPE_WEBHOOK_SECRET` | From Step 4 below |
| `BACKEND_URL` | Your Railway URL from Step 1 |

Railway restarts automatically after you save.

---

## STEP 4 — Set up Stripe webhook

1. Go to **[dashboard.stripe.com](https://dashboard.stripe.com)** → **Developers → Webhooks**
2. Click **Add endpoint**
3. Endpoint URL: `https://your-railway-url/stripe-webhook`
4. Events: select **`checkout.session.completed`**
5. Click **Add endpoint**
6. Copy the **Signing secret** (`whsec_...`) → add as `STRIPE_WEBHOOK_SECRET` in Railway

---

## STEP 5 — Update your booking form

Open `index.html` from your Netlify folder. Find this near the bottom:

```javascript
var CFG = {
  backendUrl: 'YOUR_BACKEND_URL',
  businessEmail: 'info@wh-creative.co.uk',
  stagingGuideUrl: 'YOUR_STAGING_GUIDE_PDF_URL',
};
```

Update it:

```javascript
var CFG = {
  backendUrl: 'https://your-railway-url.up.railway.app',
  businessEmail: 'info@wh-creative.co.uk',
  stagingGuideUrl: 'https://link-to-your-staging-guide.pdf',
};
```

---

## STEP 6 — Enable real-time slot checking in the form

Replace the `renderSlots` function in `index.html` with this version, which calls your backend to check which times are already booked:

```javascript
function renderSlots(dt) {
  var lbl = document.getElementById('slots-label');
  var g   = document.getElementById('slots-grid');
  document.getElementById('slots-placeholder').style.display = 'none';
  lbl.textContent = DAYS[dt.getDay()] + ' ' + dt.getDate() + ' ' + MONTHS[dt.getMonth()];
  lbl.style.display = 'block';
  g.style.display = 'grid';
  g.innerHTML = '<div style="grid-column:1/-1;text-align:center;color:#bbb;padding:20px 0;">Checking availability\u2026</div>';

  var dateStr = dt.getFullYear() + '-' +
    ('0' + (dt.getMonth()+1)).slice(-2) + '-' +
    ('0' + dt.getDate()).slice(-2);

  fetch(CFG.backendUrl + '/available-slots?date=' + dateStr)
    .then(function(r) { return r.json(); })
    .then(function(data) {
      var booked = data.bookedSlots || [];
      g.innerHTML = '';
      TIMES.forEach(function(t) {
        var taken = booked.indexOf(t) >= 0;
        var btn   = document.createElement('div');
        btn.className = 'slot-btn' + (taken ? ' booked' : '');
        btn.textContent = fmtT(t) + (taken ? ' \u2014 Taken' : '');
        if (!taken) {
          btn.onclick = function() {
            S.selectedTime = t;
            var all = document.querySelectorAll('.slot-btn');
            for (var i = 0; i < all.length; i++) all[i].classList.remove('selected');
            btn.classList.add('selected');
            document.getElementById('btn-step2').disabled = false;
          };
        }
        g.appendChild(btn);
      });
    })
    .catch(function() {
      // Fallback: show all slots if backend unreachable
      g.innerHTML = '';
      TIMES.forEach(function(t) {
        var btn = document.createElement('div');
        btn.className = 'slot-btn';
        btn.textContent = fmtT(t);
        btn.onclick = function() {
          S.selectedTime = t;
          var all = document.querySelectorAll('.slot-btn');
          for (var i = 0; i < all.length; i++) all[i].classList.remove('selected');
          btn.classList.add('selected');
          document.getElementById('btn-step2').disabled = false;
        };
        g.appendChild(btn);
      });
    });
}
```

Also add this CSS inside your `<style>` block for taken slots:

```css
.slot-btn.booked {
  background: #f5f5f5;
  color: #ccc;
  border-color: #eee;
  cursor: not-allowed;
  font-size: 12px;
}
```

Save `index.html` and re-upload the whole folder to Netlify.

---

## How it looks in your Outlook calendar

- **New booking (Stripe)** → event created with subject `📷 Photography — [address]`, marked as awaiting payment in the body
- **After seller pays** → event updated to `✅ Photography — [address]`, body updated with payment reference
- **Any event on a Monday or Tuesday** → that time slot shows as unavailable to agents in real time
- **Block time off manually** → add any event to your Outlook calendar and agents won't be able to book that slot

---

## Testing before going live

Use Stripe test mode:
- Use `sk_test_...` and update the form with `pk_test_...`  
- Test card: `4242 4242 4242 4242`, any future expiry, any CVC
- Check your Railway logs in real time: Railway dashboard → your project → **Logs**
- Switch to live keys (`sk_live_...` / `pk_live_...`) when you're ready

---

## Troubleshooting

**Calendar events not appearing** — check `MS_TENANT_ID`, `MS_CLIENT_ID`, `MS_CLIENT_SECRET` are all correct in Railway variables, and that admin consent was granted in Azure.

**Emails not sending** — make sure `Mail.Send` permission has admin consent granted. Check Railway logs for the exact error message.

**Stripe payment link not working** — confirm `STRIPE_SECRET_KEY` starts with `sk_live_` (or `sk_test_` for testing) and has no extra spaces.

**Slots all showing as available** — the `/available-slots` route may not be connecting. Check Railway logs and confirm `BACKEND_URL` in your form matches your Railway domain exactly.
