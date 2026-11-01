# RSVPs and the email list

RSVPs from daydrooler.com go into a Google Sheet owned by band@daydrooler.com
(Google Workspace; booking@, shows@, nathan@, kabir@ and grant@ are aliases of it).
`Code.gs` is a small script attached to that Sheet. Its emails come from
"Day Drooler" <band@daydrooler.com>, and replies land in the band inbox. They're sent
through band@'s own Gmail, so they also show up in its Sent folder. (Apps Script's
other mail service, MailApp, had its messages silently dropped by Gmail for this
new domain, so don't switch back to it.) It:

- adds a row for each RSVP (a repeat RSVP for the same show updates the row instead)
- emails a confirmation right away, with an "Add to calendar" link
- emails a reminder the morning before the show (around 10am Eastern)
- tells the site how many people are going to each show (counts only, never names)

The site keeps using the Partiful links until `RSVP_URL` in `index.html` is filled in.

## Set it up (about 10 minutes, once)

1. **Sign in as band@daydrooler.com** and create a new Google Sheet. Call it "Day Drooler RSVPs".
   Share it with Kabir as an editor if you both want to see the guest list.
2. In the Sheet, open **Extensions → Apps Script**.
3. Delete the starter code, paste in everything from `Code.gs`, and click **Save** (the disk icon).
   Name the project "Day Drooler RSVPs" when it asks.
4. In the function menu at the top (it says `setup` or `doGet`), pick **setup**, then click **Run**.
   Google will ask for permission:
   **Review permissions → pick band@daydrooler.com → Allow.**
   It asks to manage the Sheet, send email, and use band@'s Gmail; all three are needed.
   (If it shows "Google hasn't verified this app", click **Advanced → Go to Day Drooler RSVPs (unsafe)** first.
   That warning appears for any script Google hasn't reviewed; this one is yours.)
   When it finishes, the Sheet has an **RSVPs** tab with column headers.
5. Click **Deploy → New deployment**. Click the gear next to "Select type" and choose **Web app**. Set:
   - **Execute as:** Me
   - **Who has access:** Anyone

   Click **Deploy** and copy the **Web app URL** (it ends in `/exec`).
6. Put that URL in `index.html` (`const RSVP_URL = "…";`), then commit and push.
   The RSVP buttons now open the form on the site.
7. **Test it:** RSVP to a show with your own email. You should see a row in the Sheet,
   get the confirmation email, and see "1 going" on the site. Delete the row afterwards.

## Changing the script later

Edit the code in the Apps Script editor, then **Deploy → Manage deployments → pencil icon →
Version: New version → Deploy**. Saving alone doesn't update the live web app.
Always edit the existing deployment rather than making a new one, so the web app URL
in `index.html` stays the same. If a change needs new permissions, run `setup` once
from the editor first so Google can ask for them.
Keep this file in sync so the repo has the current version.

## Good to know

- **The script can send about 1,500 emails a day** from a Workspace account. If a busy day goes over,
  the extra confirmations go out at the next morning's run. The Confirmation column
  shows "sent", "waiting (daily limit)", or why one failed.
- **The email list** is every row with "yes" in the Email list column. To email them, filter
  on that column and copy the addresses into BCC. Once the list is in the hundreds, import
  it into a newsletter service (Kit, Buttondown) so unsubscribes are handled for you.
- **To remove someone**, delete their row. Counts on the site update within a minute.
- **The going count is public** (for example "3 going"); names and emails are never sent to the site.
- This folder is public like the rest of the repo. The script holds no passwords or keys,
  so that's fine, but never paste the Sheet's contents in here.
