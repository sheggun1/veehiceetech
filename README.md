# Veehicee Telecommunications and Services

Frontend: GitHub Pages (`index.html`, `marketplace.html`). Backend: Google Apps Script (`Code.gs`). Database: Google Sheets. Media: Google Drive.

## Important security notes

Keep payment secret keys out of GitHub and frontend code. Store them only in Apps Script Project Settings > Script Properties:

- `VEEHICEE_PAYSTACK_SECRET`
- `VEEHICEE_FLW_SECRET_KEY`
- `VEEHICEE_NOTIFY_EMAIL` (optional)

## Deploying the Apps Script upgrade

1. Back up the existing Apps Script project and current `Code.gs`.
2. Copy the latest `Code.gs` from the `main` branch into the connected Apps Script project and save.
3. Run `runVeehiceeSetup()` from the Apps Script editor. Use the administrator email already configured in Script Properties. This setup adds missing sheets and headers; it does not clear existing rows.
4. Review the new sheets in the connected Google Sheet.
5. Deploy > Manage deployments > Edit the existing web app > select **New version** > Deploy. Keep **Execute as Me** and the existing web app URL.
6. For automatic abandoned-cart reminder emails, run `setupVeehiceeCartReminderTrigger()` once from the Apps Script editor and approve the requested permissions. It schedules a daily check at approximately 10:00 in the Apps Script project time zone.
7. Confirm that the existing `API_URL` in both frontend files still points to the current `/exec` web app URL.

## Added in the business-management upgrade

- Admin product cost-price and wholesale-price fields. Wholesale price is applied to a line quantity of five or more.
- Expense register, sales summary, estimated profit, and monthly revenue summary.
- Customer directory with order totals and CSV exports for customers, orders, and expenses.
- Customer accounts, order history, printable invoices, wishlists, product comparisons, and saved carts.
- Coupon creation and validation, plus loyalty points (1 point = ₦1) with points redemption on checkout.
- Verified-purchase product reviews with administrator approval and visible ratings.
- Customer price-drop email alerts when an administrator lowers a product price.
- Low-stock email report and abandoned-cart reminder tools.
- Optional daily abandoned-cart reminder trigger.
- WhatsApp click-to-chat links for order follow-up (not a WhatsApp Business API integration).
- Homepage and marketplace SEO/social metadata, canonical URLs, `robots.txt`, and `sitemap.xml`.

## Operational notes and limitations

- Enter accurate product cost prices before relying on estimated profit. Older orders created before cost-price capture may not contain product costs.
- Set actual stock quantities, prices, wholesale prices, shipping fees, business contact details, and payment settings before promoting products.
- The WhatsApp feature opens a WhatsApp conversation; automatic WhatsApp templates/messages require a WhatsApp Business API provider and credentials.
- Abandoned-cart emails require customers to sign in and choose **Save my cart**. The daily trigger is optional and must be enabled once by the Apps Script owner.
- Payment credentials, Apps Script deployment, Google Sheets/Drive permissions, email permissions, and real payment sandbox flows still need to be configured and tested. This repository update does not itself deploy Apps Script or test live transactions.
- Refund initiation, automated loyalty reversal for refunds, and external WhatsApp API messaging are not included in this release.
- Do not advertise online payments as active until Paystack/Flutterwave sandbox and live deployment tests pass.
