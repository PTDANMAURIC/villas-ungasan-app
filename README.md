# Villas Ungasan — operations app

One link on everyone's phone. Owners sign in to scan invoices, see expenses and pay workers. Workers sign in to check in and out with a live photo and see their to-do list.

| Who | Can do |
|---|---|
| **Admin** (Leon) | Everything, including adding and removing people and resetting PINs |
| **Owners** (Mauri 131, Dani 132 + 133, Ric 134) | See, add, edit and delete expenses for **all villas**, split an invoice between villas, see workers' days and salaries, give tasks |
| **Workers** | Check in and out (live camera only, no gallery; GPS is saved), see and tick off their tasks, see their days worked. English/Bahasa toggle |

**Where the data lives (your Google account):**
- Google Sheet **"Villas Ungasan – App"** with tabs *Expenses*, *Attendance*, *Tasks*, *Users* and *Summary* (totals per villa per month, updated automatically)
- Drive folder **"Villas Ungasan – App"**
  - `Invoices / Villa 131 – Mauri`, `Villa 132 – Dani`, … (one folder per villa, files named `2026-11-05_131_Vendor_1250000.jpg`)
  - `Invoices / Shared between villas` (invoices split across several villas, e.g. `2026-11-05_132-133-134_Vendor_1000000.jpg`)
  - `Attendance photos / 2026-11 / …`

---

## Setup (about 15 minutes, once)

### 1. Create the backend (Google Apps Script)
1. Go to **script.google.com**, signed in with the Google account that should own the data. Click **New project** and name it *Villas Ungasan API*.
2. Delete what's in `Code.gs`, then paste in the whole contents of **`google-apps-script/Code.gs`**.
3. Optional: at the top, in `SETUP`, change names or add workers. You can also add workers later in the app.
4. Choose the function **`setup`** in the toolbar dropdown, then click **▶ Run**. Google will ask for permission (Drive and Sheets). Click *Advanced → Go to Villas Ungasan API (unsafe)*. This warning appears because it's your own unpublished script.
5. Open **View → Execution log** to see the starting **PIN** for Leon, Mauri, Dani and Ric. Hand them out privately. Everyone can change their PIN in the app (⋮ menu).

### 2. Switch on automatic invoice reading
In the Apps Script editor, open **Project Settings (⚙) → Script properties → Add property** and add **one** of these:
- `ANTHROPIC_API_KEY` = a Claude API key from console.anthropic.com. This reads Indonesian handwritten *nota* very well and costs about 1–2 US cents per invoice.
- **or** `GEMINI_API_KEY` = a key from aistudio.google.com/apikey. It has a free tier.

Without a key the app still works, but owners type the invoice details by hand. The photo is still saved.

### 3. Publish the backend
**Deploy → New deployment → ⚙ Select type: Web app**
- Execute as: **Me**
- Who has access: **Anyone**

Click **Deploy** and copy the **Web app URL** (it ends in `/exec`).

> "Anyone" only means the address can be reached. Nobody can see or change anything without a name + PIN.

### 4. Publish the app page
The app is the single file **`index.html`** (in the root of this repository). Put your `/exec` URL into it:
```js
const API_URL = 'https://script.google.com/macros/s/XXXX/exec';
```
Then host it on **GitHub Pages**: repository → *Settings → Pages → Deploy from branch → main / root*. The link looks like `https://<your-username>.github.io/villas-ungasan-app/`.
(Netlify Drop or Cloudflare Pages work too. Any HTTPS host is fine, because the phone camera only works over HTTPS.)

### 5. Put it on everyone's phone
Send the link. On the phone, open it and choose **Add to Home Screen**. It then opens like an app.
The first time a worker checks in, the phone asks to allow the **camera** and **location**. They must tap **Allow**.

---

## Daily use

**Owners: scan an invoice**
1. Tap **Scan invoice**, then **Take photo** (or upload a photo or PDF from WhatsApp). Printed and handwritten *nota* both work.
2. After about 10 seconds the form fills in automatically: date, shop, what was bought, category and total.
3. **Choose the villa.** To share the cost, tap several villas (or **All 4**):
   - **Split equally** divides it for you in whole rupiah. For example, Rp 1.000.000 across 3 villas becomes 333.334 / 333.333 / 333.333, and you see this before saving.
   - **Custom amounts** lets you type each villa's share. The app won't save until the shares add up exactly to the invoice total.
4. **Check every field, especially the amount**, then tap **Save**.

In the *Expenses* sheet, a shared invoice becomes one row per villa. Each row has that villa's share in **amount**, the full invoice in **invoiceTotal**, and the split written out in **split**. All the rows point to the same invoice photo. Per-villa totals in *Summary* therefore stay correct.
In the app, **All villas** shows each invoice once at its full amount. Tapping a villa shows only its share ("Rp 333.333 of Rp 1.000.000").
Any owner or the admin can edit or delete any invoice. The *Expenses* sheet keeps who added each one. Editing or deleting a shared invoice changes it for every villa.

**Workers**
- Arrive: **Check in**, take a selfie on site, then **Confirm**. There's no villa to choose, because all four villas are on the same site.
- Leave: **Check out** with a selfie.
- Tick tasks when they're done.

**End of month: pay salaries**
Open **Workers**, then pick the month. You see each worker's **days worked** (any check-in = 1 day), hours, and **salary = days × daily rate**. Days without a check-out are flagged in orange.
- The total is split **equally between the 4 villas**. Tap **Add salaries to expenses** to save it as one *Staff salaries* expense, so each villa's monthly total includes its share. It opens the split screen with an equal four-way split. In a month when one villa pays more or less, choose **Custom amounts** and type each villa's share. If you tap it again for the same month, it updates that expense instead of adding a second one.
- Tap a worker to see each day with photos and location, fix a forgotten check-out, or **add a day manually** (for example, phone broken).
- **Download payroll CSV** for your records.

**Team (admin only):** add workers (name, daily rate, PIN), switch people off when they leave, reset PINs.

## Good to know
- **Forgotten PIN:** go to *Team → person → New PIN*. Or, in Apps Script, run `resetPin("Name")`.
- **5 wrong PINs** lock that name for 15 minutes.
- **Times** are Bali time (WITA). Amounts are saved as positive numbers. The AI never saves on its own: a person always confirms.
- **Changing the backend code later:** paste the new code, then *Deploy → Manage deployments → ✏ Edit → Version: New version → Deploy*. The URL stays the same.
- **Why the app is hosted outside Google:** Google blocks the live camera inside Apps Script pages, so the screens are hosted separately and only talk to Google for data.
