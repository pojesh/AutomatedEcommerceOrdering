# UiPath Automated E-Commerce Ordering Process

This repository contains a UiPath automation project designed to process product orders from an Excel file received via email. The robot performs end-to-end web automation on the "Swag Labs" e-commerce test site ([saucedemo.com](https://www.saucedemo.com)) to place orders, then updates and replies to the original email with the results.

---

## Features

### ✉️ Email Automation
- Monitors a Gmail inbox for emails with a specific subject line ("Product Order Request").

### 📄 Data-Driven Processing
- Downloads, reads, and iterates through an Excel file containing product and customer data.

### 🌐 Web Automation
- Logs into saucedemo.com.
- Selects products and performs checkout.
- Extracts order details.

### 🔗 Dynamic UI Interaction
- Uses dynamic selectors to identify and click elements based on data from the Excel file.

### ⚖️ Data Manipulation
- Extracts and cleans pricing info.
- Updates Excel with order status, prices, and screenshot paths.

### 📈 Reporting
- Takes screenshots of order confirmations.
- Sends reply email with the updated Excel file.

---

## Process Workflow

### 1. Monitor Gmail
- Checks Gmail for unread emails with subject containing "Product Order Request".
- Logs number of matching emails.

### 2. Process Each Email
- Downloads attachments.
- Moves & renames the file to `/Data/ProductOrderDetails.xlsx`.

### 3. Read and Validate Excel
- Loads data from `ProductOrderDetails.xlsx` into `dt_ExcelData`.
- Checks if it contains valid rows.

### 4. Perform Web Automation
- Opens `saucedemo.com` in Edge.
- Logs in with `standard_user / secret_sauce`.

#### For Each Product:
- Extracts `ProductName` from current row.
- Clicks product link using dynamic selector.
- Clicks "Add to cart".
- Proceeds to cart → checkout.
- Fills First Name, Last Name, and Postal Code.
- Clicks "Continue".
- Extracts `ItemTotal`, `Tax`, and `TotalCost`.
- Cleans values and updates `dt_ExcelData`.
- Clicks "Finish".
- Takes a screenshot, saves to `/Data/Screenshots/Order_<ProductName>.png`.
- Updates row with "Order Placed" and screenshot path.
- Clicks "Back Home".

### 5. Finalize and Reply
- Saves updated Excel file to `/Data/ProductOrderDetailsUpdated.xlsx`.
- Sends a reply email with the updated file attached.

---

## Prerequisites

- UiPath Studio (latest version)

### Required Packages:
- UiPath.Excel.Activities  
- UiPath.GSuite.Activities  
- UiPath.Mail.Activities  
- UiPath.System.Activities  
- UiPath.UIAutomationNext.Activities

---

## Configuration

### ⚙️ GSuite Connection
- Configure in UiPath Studio with necessary credentials.

### 📁 Folder Structure
```
/Data
/Data/Screenshots
```
- Ensure folders exist or are created before execution.

### 📧 Input Email Format
- Subject: `Product Order Request`
- Attachment: Excel file with sheet named `Sheet1`

---

## How to Run

1. Open project in UiPath Studio.
2. Configure GSuite connection in Connectors.
3. Send email to target Gmail account (with proper subject & attachment).
4. Run `Main.xaml`.

---

## Key Variables

| Variable Name         | Data Type                      | Scope          | Description |
|-----------------------|--------------------------------|----------------|-------------|
| `dt_ExcelData`        | `System.Data.DataTable`        | Main Sequence  | Product/customer data from Excel |
| `EmailList`           | `List(Of GmailMessage)`        | Sequence Top   | Emails fetched from Gmail |
| `currentGmailMessage` | `GmailMessage`                 | For Each       | Current email being processed |
| `DownloadedAttachment`| `GmailAttachmentLocalItem[]`   | Sequence Body  | Downloaded attachments |
| `CurrentRow`          | `System.Data.DataRow`          | For Each Row   | Current product row |
| `ProductName`         | `String`                       | Sequence Body  | Name of product from Excel row |
| `ItemTotal`           | `String`                       | Sequence       | Cleaned item price |
| `Tax`                 | `String`                       | Sequence       | Cleaned tax value |
| `TotalCost`           | `String`                       | Sequence       | Cleaned total cost |
| `ScreenshotPath`      | `String`                       | Sequence       | Path to screenshot file |
| `UpdatedExcelPath`    | `String`                       | Sequence       | Path of final updated Excel file |

---

