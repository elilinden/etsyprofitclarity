# Etsy Profit Clarity

A lightweight browser app for Etsy sellers to upload Etsy payment-account CSV or Excel exports, review sales and Etsy deductions, enter product and fulfillment costs, and download a clean profit report.

## Features

- Upload multiple `.csv`, `.xlsx`, or `.xls` Etsy statement files.
- Group transactions by Etsy order number.
- Track gross sales, Etsy deductions, seller-entered costs, and estimated profit.
- Add default, product-specific, and order-specific costs.
- Download a multi-tab Excel report or CSV summary.

## Run Locally

This is a static site. Open `index.html` directly, or serve the folder locally:

```bash
python3 -m http.server 5173
```

Then visit `http://127.0.0.1:5173/`.
