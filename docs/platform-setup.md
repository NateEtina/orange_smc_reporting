# Platform Setup — Telegram & WhatsApp Cloud API

> Step-by-step configuration for connecting platforms to the bot backend.
>
> **Caveat:** Meta's dashboard UI changes frequently. Treat exact menu labels as
> approximate; the API contracts described here are stable.

---

## A. Telegram (~30 minutes, no approval queue)

### 1. Create the bot
1. Open Telegram, search **@BotFather**, start a chat.
2. Send **`/newbot`**.
3. Pick a **display name** (anything) and a **username** — the username *must*
   end in `bot` (e.g. `acme_support_bot`).
4. BotFather returns your **HTTP API token**: `7123456789:AAH...`
   - Shown once. Recover via `/mybots` → *API Token* → *Revoke*.
   - One token **per tenant bot account** in the SaaS.

### 2. Configure the bot's surface (all in BotFather)
- **`/setdescription`** — shown before people start the bot.
- **`/setabouttext`** — profile "about" line.
- **`/setuserpic`** — avatar.
- **`/setcommands`** — autocomplete menu:
  ```
   start - Begin using the bot
   stop - Pause auto-replies
   help - Get support
   status - Post to channel status
  ```
- **`/setprivacy` → `Disable`** — **critical.** By default bots in groups only
  see replies/mentions. Disable privacy to read all group messages.

### 3. Choose your update transport
Exactly one of two — Telegram rejects a webhook if `getUpdates` has pending updates.

| Mode | Use when | How |
|---|---|---|
| **Webhook** (prod) | Public HTTPS URL | `setWebhook` |
| **Long polling** (dev) | Local dev, no public URL | `GET /bot<token>/getUpdates?timeout=30` |

**Webhook:**
```bash
curl "https://api.telegram.org/bot<token>/setWebhook" \
  -d "url=https://api.yourapp.com/webhooks/telegram" \
  -d "secret_token=whsec_<random-64>" \
  -d 'allowed_updates=["message","edited_message","channel_post","callback_query"]'
```

Key parameters:
- **`secret_token`** → Telegram sends it as header
  **`X-Telegram-Bot-Api-Secret-Token`**. Reject on mismatch. Telegram signs
  nothing, so this is your only verification.
- **`allowed_updates`** — omitting it queues unhandled update types forever.
  Always set it.
- **`drop_pending_updates=true`** — clears backlog after redeploy.
- **HTTPS is mandatory** for webhooks.

Verify:
```bash
curl "https://api.telegram.org/bot<token>/getWebhookInfo"
```
`last_error_message` and `pending_update_count` are your first debugging tool.

### 4. Handle the payload
JSON shaped as `{"update_id": 987654321, "message": {...}}` — `update_id` is
strictly increasing per bot → **idempotency key**.

Respond `200` immediately (empty body) and enqueue. Fields to normalize:
- `message.from.id` → sender
- `message.chat.id` → destination (negative for groups; `@channelusername` also works)
- `message.text`, `message.photo[]`, `message.document` → content
- `channel_post` → for channel/status posting

### 5. First send
```bash
curl -X POST "https://api.telegram.org/bot<token>/sendMessage" \
  -d "chat_id=<CHAT_ID>" -d "text=Hello from the bot"
```

### 6. Channel / status posting
1. Create your channel.
2. Add the bot as an **administrator** with *Post messages* permission.
3. Post with `chat_id=@yourchannel` or the numeric `-100...` ID.

### Telegram constraints that shape the design

| Constraint | Impact |
|---|---|
| **A bot cannot cold-message someone.** User must `/start` first. | Kills broadcast-to-strangers; auto-reply is inbound-initiated only. |
| **~48-hour activity window** after the user's last interaction. | Telegram's version of WhatsApp's 24h rule — rules engine must branch on it. |
| **Rate limit ~30 msgs/sec overall**, ~20/min per group. | Enforce per-tenant in the outbound queue. |
| **No App Review, no verification, no fees.** | Build second — validates the adapter interface with zero friction. |

---

## B. WhatsApp Business Cloud API

Two phases: **test in ~30 min**, then **production (1–3 weeks, mostly review queue)**.

### Phase 1 — Test environment (30 min)

**Prerequisites**
- Meta developer account → https://developers.facebook.com
- A **Meta Business Account** (Business Manager) — required for production, create now.
- 🚩 **A phone number not registered on any WhatsApp app.** Registering it on
  Cloud API detaches it from the app. Use a spare/VOIP number for testing.
- A public HTTPS endpoint.

**Steps**

1. **developers.facebook.com → My Apps → Create App**
   - App type: **Business** (not Consumer).
   - Name it, attach your Meta Business Account.

2. **Add the WhatsApp product.**
   (Meta periodically ships a *"Start with WhatsApp"* guided flow — same destination.)

3. **Grab test credentials** from *WhatsApp → API Setup*:
   - **Temporary access token** — ⚠️ **expires in 24 hours.** Dev only.
   - **From phone number ID** — Meta provides a test number; this ID is what you call.
   - **Test recipient** — up to 5 numbers you may message while testing.

4. **Get a permanent token** now:
   - *Meta Business Suite → Settings → Business settings → Users → System users → Add*
   - Role: **Admin** → *Generate new token* → select your app
   - Scopes: **`whatsapp_business_messaging`** + **`whatsapp_business_management`**
   - Copy it — **shown only once.**

   Store encrypted per-tenant; rotate every **~60 days**.

5. **Register your webhook** — *WhatsApp → Configuration → Webhook*:
   - **Callback URL**: `https://api.yourapp.com/webhooks/whatsapp`
   - **Verify token**: a static string you invent. Store as env var.
   - **Verify and save** → Meta issues a GET:
     ```
     GET /webhooks/whatsapp?hub.mode=subscribe
                          &hub.verify_token=<YOUR_TOKEN>
                          &hub.challenge=<RANDOM_STRING>
     ```
     Respond **`200` with body = `hub.challenge`, verbatim**:
     ```python
     @router.get("/webhooks/whatsapp")
     async def verify(hub_mode: str, hub_verify_token: str, hub_challenge: str):
         if hub_mode == "subscribe" and hub_verify_token == settings.WA_VERIFY_TOKEN:
             return PlainTextResponse(hub_challenge)
         raise HTTPException(403)
     ```
   - **Subscribe to the `messages` field.** (Later: `template_status`, etc.)

6. **Verify the signature** on every POST. Meta signs the **raw body** with your
   **App Secret**:
   ```
   X-Hub-Signature-256: sha256=<hex hmac>
   ```
   ```python
   expected = "sha256=" + hmac.new(
       APP_SECRET.encode(), raw_body, hashlib.sha256
   ).hexdigest()
   if not hmac.compare_digest(expected, header_value): raise HTTPException(401)
   ```
   ⚠️ Needs the **unparsed** `bytes` — read `await request.body()` *before*
   FastAPI parses JSON.

7. **Send your first message:**
   ```bash
   curl -X POST "https://graph.facebook.com/v2X.0/<PHONE_NUMBER_ID>/messages" \
     -H "Authorization: Bearer <TOKEN>" \
     -H "Content-Type: application/json" \
     -d '{
       "messaging_product": "whatsapp",
       "to": "<RECIPIENT_E164>",
       "type": "text",
       "text": {"body": "Hello from the bot", "preview_url": false}
     }'
   ```
   🚩 **Pin the Graph API version** (copy from dashboard, e.g. `v2X.0`) in
   config. Auto-following `latest` lets Meta break you silently on release day.

8. **Receive a message** — inbound body:
   ```json
   {"object":"whatsapp_business_account",
    "entry":[{"changes":[{"value":{"messages":[
      {"from":"1555...","id":"wamid.AB...","timestamp":"...","type":"text",
       "text":{"body":"hi"}}]}]}]}}
   ```
   - **`id` (`wamid`)** is globally unique → idempotency key.
   - Respond `200` in **under ~5 seconds** or Meta marks it failed and **retries
     with the same `wamid`**.
   - May deliver **multiple messages per POST** — loop over them.

### Phase 2 — Production access (1–3 weeks)

1. **Business Verification** — *Business Settings → Security Center*. Needs
   utility bills/corporate docs. **Start immediately; it's the long pole.**

2. **Register your actual phone number** — *WhatsApp → Phone numbers*:
   - Set business display name + about.
   - **6-digit code** arrives by SMS/voice (or in-app if the number is already on
     WhatsApp).
   - Display name must not look generic or brand-mismatched.

3. **Move test → live**, request **tier upgrade**:
   - Default: **Tier 1 = 1,000 unique recipients / 24h**.
   - 10K → 100K → unlimited unlock with **good quality rating + 2FA +
     business verification**.
   - Watch *Phone numbers → Quality rating*. Yellow/Red throttles you.

4. **App Review** for `whatsapp_business_messaging` +
   `whatsapp_business_management`, plus webhook field permissions. Record a
   screencast. This is the 1–2 week wait.

5. **Message templates** — *WhatsApp Manager → Message templates* → approval
   (minutes to days):
   ```json
   {"messaging_product":"whatsapp","to":"...","type":"template",
    "template":{"name":"order_update","language":{"code":"en_US"},
      "components":[{"type":"body","parameters":[{"type":"text","text":"#1234"}]}]}}
   ```

### WhatsApp constraints that shape the design

| Constraint | Impact |
|---|---|
| **24-hour session window.** Inside → free-form. Outside → **approved templates only.** | ⚠️ Biggest constraint on the auto-reply product. Rules engine must branch on "inside window?" and template-ize anything outside. |
| **No Status posting via Cloud API.** | "Post on Status" **requires the Node/Baileys sidecar.** Cloud API cannot do it. |
| **Rate limits by tier** (1K/10K/100K/unlimited per 24h). | Per-tenant rate limiter in outbound queue. |
| **Template approval + business verification are async.** | Allow `connected but not yet approved` state; don't block the UI. |
| **Inbound media URLs expire quickly.** | Download to own storage inside the worker, immediately. Never store the URL. |
| **Cost per conversation** by category (marketing / utility / service). | Log conversation metadata per message → per-tenant cost dashboard. |

---

## C. Mapping to the codebase

**Endpoints**
```
GET  /webhooks/whatsapp     → Meta challenge/verify
POST /webhooks/whatsapp     → verify X-Hub-Signature-256 → persist → enqueue
GET  /webhooks/telegram     → health/optional
POST /webhooks/telegram     → verify X-Telegram-Bot-Api-Secret-Token → persist → enqueue
```

**Config (env, dev only)**
```bash
WA_VERIFY_TOKEN=...
WA_APP_SECRET=...            # signature check
WA_SYSTEM_USER_TOKEN=...     # dev token; prod lives encrypted in DB
WA_GRAPH_VERSION=v2X.0
TELEGRAM_WEBHOOK_SECRET=...
```
🚩 In production, **platform tokens are per-tenant and encrypted in Postgres**,
not env vars. Env vars hold only app secrets, verify tokens, and encryption keys.

**Tables**
```sql
platform_accounts
  id, tenant_id, platform, external_id, phone_number_id,
  chat_id, token_encrypted, status(pending|active|revoked),
  quality_tier, webhook_fields[], created_at

webhook_events
  id, platform, external_event_id UNIQUE, raw jsonb,
  received_at, processed_at, status
```
`UNIQUE(platform, external_event_id)` (`wamid` / `update_id`) is what makes
webhook retries safe. Add on day one — retrofitting dedupe onto a live bot
means users get double replies.

---

## D. Local development

Neither platform reaches `localhost`. Use a tunnel:
```powershell
cloudflared tunnel --url http://localhost:8000
# or
ngrok http 8000
```
Point the webhook at the issued `https://xxx.trycloudflare.com` URL.

---

## Suggested sequence

```
1. Telegram via long-polling     → proves adapter interface, zero friction
2. Telegram via webhook + tunnel → proves verify/persist/enqueue path
3. WhatsApp test token           → inbound/outbound round trip
4. WhatsApp production           → start business verification on day 1
                                   (gates everything, takes weeks)
```
