SONG REQUESTS — INSTALLATION

This package is an addition to your existing Apps Script project for spreadsheet:
https://docs.google.com/spreadsheets/d/1pv3g2oRsBX6JuWqVTNfz-GIYV3TTTtckD1YURVSc7OU/edit

1. Open the spreadsheet > Extensions > Apps Script.
2. KEEP your existing PDF functions, including its onOpen() function.
3. Add a new Script file called SongRequests and paste SongRequests.gs.
4. If you already have a doGet() function from the previous song-search prototype, DELETE that older doGet() function. Keep only the new one. If you have an older Index HTML file, you may leave it unused.
5. Add two HTML files named Public and Admin. Paste Public.html and Admin.html respectively.
6. Save. Select setupSongRequests in Apps Script, click Run and complete the prompted 16+-character unique admin password. Authorize requested Google permissions. This creates Sessions, Requests and Votes tabs. Never share your admin password.
7. Deploy > New deployment > Web app > Execute as Me > Who has access: Anyone. Deploy and copy your /exec link. Visitors use /exec ; your admin address is /exec?admin=1 (password still required).
8. After editing source later, use Deploy > Manage deployments > Edit > New version > Deploy. A code change alone does not update an existing /exec deployment.
9. Scan public URL with an iPhone and test: start session (admin), search 3+ letters, request, vote, queue, played, close requests, close session. Save public URL as a QR code only after testing.

SESSION WORKFLOW
- OPEN: accept requests and votes.
- NO_REQUESTS: public can browse queue, but cannot add requests or votes. Admin can still reorder, park, mark played.
- CLOSED: no session active; archive remains in Sheets and in a new Drive folder.
- In admin, enter a positive integer in Order and click Queue / Move. Smaller numbers play earlier. Set a larger number to move a song later. Order is ignored when status becomes PLYD.
- A song can be requested once per evening. The requester can cancel it while REQU; cancelling makes its existing votes inapplicable. Others may request it again thereafter.
- All requests/votes are saved immediately in Google Sheets and archived as CSV on Close Session.

SECURITY AND LIMITS
- This is an initial community-use app, not a security-audited product.
- Admin actions are protected by a server-side long-password check and an expiring admin token; ?admin=1 alone grants no admin access.
- Browser ID / localStorage limits normal double-voting only. Users can deliberately create more identities with other browsers, private mode or cleared data. It does NOT enforce one physical phone, and should not be advertised as such.
- For stronger voter identity controls, integrate a trusted identity provider or phone verification via a separate service; Apps Script alone does not have physical-device attestation.
- Every public visitor can read song and artist names and request/queue statuses. No private Playlist columns are shared via the public API.
- The public web app runs with the sheet owner's Drive and Sheet permissions. Do not add unauthenticated functions that expose private data or modify status.
- This app does not add links to the existing PDF menu and does not replace the original PDF scripts.
