# Bank Statement Converter

A browser-local converter for Bank of Communications (交通银行) XLS/XLSX statements into a WeChat Pay-style XLSX workbook.

## Privacy

All statement parsing and conversion happens in your browser. The source statement is not uploaded to GitHub or any application server. The page loads the SheetJS parser from its public CDN; no transaction data is sent to that CDN. Avoid using a shared/public computer and close the page when finished.

## Use

1. Open the GitHub Pages site.
2. Select a Bank of Communications `.xls` or `.xlsx` statement.
3. Check the detected columns and preview.
4. Export the WeChat-compatible `.xlsx` file.

The source format is expected to contain these columns: 记账日期、交易时间、交易地点、交易方式、支出金额、收入金额、余额、对方户名、对方开户行. Header detection scans the first 30 rows.

The output workbook uses sheet `Sheet1`, places the 11 WeChat-style column headers on row 18, and preserves all 11 columns: 交易时间、交易类型、交易对方、商品、收/支、金额(元)、支付方式、当前状态、交易单号、商户单号、备注.

Fields not present in the bank statement (交易单号、商户单号) remain blank. The tool does not fabricate transaction identifiers. Bank-posted records are labeled `已完成`; verify this mapping against your own WeChat import requirements before relying on it.

## Scope

- Bank of Communications XLS/XLSX only.
- WeChat-style XLSX export only; Alipay is out of scope for this version.
- No server-side upload, storage, or processing.

## Deployment

GitHub Actions deploys the static site to GitHub Pages on pushes to `main`. Enable Pages with **Build and deployment → Source: GitHub Actions** if it is not already enabled.