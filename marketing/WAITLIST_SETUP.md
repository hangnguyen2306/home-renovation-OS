# Waitlist + free checklist: setup (do this now, before the BV exists)

**Goal:** people who find your pins download the free **Renovation Money Checklist** and join your email list. On launch day you email them the launch discount. No sales happen yet, so there's no registration issue. Confirm this with your accountant.

## 1. Choose a free email tool (not Gumroad yet)
Use a free email-marketing tool with **landing pages + automated emails**, for example **MailerLite** (free plan) or **Kit / ConvertKit** (free plan). Open **Gumroad later, in the BV's name**.

## 2. Build the signup page (about 30 minutes)
- **Headline:** "FREE: The Renovation Money Checklist"
- **Sub-headline:** "15 hidden costs to check in every quote, payment rules that protect you, and the documents you must keep — 5 printable pages."
- **Image:** `marketing/leadmagnet_cover.png`
- **Form fields:** first name (optional) + email
- **Consent checkbox** (required in the EU, must *not* be pre-ticked):
  "Yes, send me the checklist and occasional emails about renovation planning and the launch of Renovation OS. I can unsubscribe at any time."
- Link to a short **privacy notice**: who you are, what you collect (email and name), why (checklist + launch news), the tool you use, and how to unsubscribe or ask for deletion. Most tools have a template.
- Turn on **double opt-in** (confirmation email). It's GDPR-friendly and keeps the list clean.

## 3. Automated welcome email
Subject: *Your Renovation Money Checklist 📋*
```
Hi {first name},

Here is your free Renovation Money Checklist: [download link]

Print it and take it to every contractor meeting — tick every box before you sign a quote.

I'm building Renovation OS, an Excel planner that does all of this for you: it checks
hidden costs in quotes, warns you when you pay a contractor ahead of the work and ranks
your to-dos every Monday. You're on the list, so you'll get the launch discount first.

One question: what is the hardest part of your renovation right now? Just hit reply —
I read every answer.

Hang
```
Upload `Renovation_Money_Checklist.pdf` to the tool, or to a cloud folder shared as "anyone with the link can view", and paste the download link.

The replies to "hardest part" are gold: they tell you what to post about and what to add to the product.

## 4. Connect the pins
- Put the signup-page URL in `WAITLIST_LINK` in `marketing/make_pins.py`, then run `python3 marketing/make_pins.py --waitlist`.
- Upload and schedule the 20 pins from `marketing/pins_waitlist/` (see `pins/HOW_TO_POST.md`), starting with **19** and **20**, which promote the checklist directly.

## 5. On launch day (BV + Gumroad ready)
1. Email the list: "Renovation OS is live — your launch code LAUNCH20 is valid for 7 days" + link.
2. **Edit the link of your pins** to the Gumroad product. You can change the destination URL of your own pins in Pinterest: open the pin ▸ ⋯ ▸ Edit Pin. Start with the pins that get the most clicks.
3. New pins from then on come from `python3 marketing/make_pins.py` (product link).
4. Keep the free checklist running: it keeps growing your list after launch too.
