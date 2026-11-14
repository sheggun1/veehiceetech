# Veehicee website upgrade

Backend: Google Apps Script (`Code.gs`); frontend: `index.html`; database: Google Sheets; media: Google Drive.

## Important
Keep payment secret keys out of GitHub and frontend code. Store them only in Apps Script Project Settings > Script Properties:
- `VEEHICEE_PAYSTACK_SECRET`
- `VEEHICEE_FLW_SECRET_KEY`
- `VEEHICEE_NOTIFY_EMAIL` (optional)

## Deployment
1. Back up the existing Apps Script project and Code.gs.
2. Copy the Code.gs from this branch into the connected Apps Script project.
3. Run setupVeehicee with the administrator email and password you already use. The setup adds missing sheets and columns without clearing existing rows.
4. Review the spreadsheet and Drive folder.
5. Add provider keys to Script Properties when merchant sandbox accounts are ready.
6. Deploy > Manage deployments > Edit > New version. Keep Execute as Me. If you create a new deployment, update API_URL in index.html.
7. Test checkout, delivery, media upload and payment sandbox flows before accepting live orders.

## Included
Admin login throttling and password migration to PBKDF2-HMAC-SHA256; admin password changes; new Sheets tables; portfolio and product management; Drive image uploads; customer accounts and guest order tracking; cart and checkout; stock reservation; Paystack and Flutterwave server-side verification callbacks; manual bank transfer confirmation; delivery status; email notifications; business and SEO settings; inventory report.

## Launch limitations
The GitHub update does not deploy Apps Script or configure the Google Sheet, Drive, or payment merchant accounts. Live integration has not been tested. Set real delivery fees and bank details. Refund initiation is not automated. Do not merge to main or advertise online payments as active until sandbox and live deployment tests pass.