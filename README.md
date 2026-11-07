# gặp — gapevents.vn

One-page website for gặp, social dining in Saigon. Plain HTML + images, hosted on Netlify (publishes from the `main` branch).

- `index.html` — the page
- `images/` — logos, hero, story/host photos, and the carousel (`gallery/NN-thumb.jpg` thumbnails, `gallery/NN.jpg` full size for the lightbox)

## Event pages (free events with sign-up)

- `events/_template.html` — copy to `events/<event-id>.html` for each event (served at `gapevents.vn/events/<event-id>`). Every editable bit is marked `EDIT`.
- `events/event.css`, `events/register.js` — shared by all event pages. `GAP_ENDPOINT` at the top of `register.js` is the Apps Script web app URL.
- `apps-script/Code.gs` — the sign-up backend. It lives in the "gặp — Event registrations" Google Sheet (Extensions → Apps Script). It saves sign-ups to the Registrations tab, applies capacity from the Events tab, puts overflow on a waitlist and emails confirmations from hello@gapevents.vn.
- QR check-in: every sign-up gets a ticket code (Registrations column M) and confirmed guests get it as a QR in their email. Staff open `gapevents.vn/checkin` on a phone, enter the staff PIN (Apps Script → Project Settings → Script properties → `STAFF_PIN`) and scan guests in. It ticks "Checked in" and stamps the time (column N). "Find a guest" covers anyone without their QR. Add `?event=<event-id>` to the link to skip the event picker.
- Capacity is set only in the Sheet. The page never shows numbers, just: spots are limited → half the spots are gone → almost fully booked → fully booked · join the waitlist.
