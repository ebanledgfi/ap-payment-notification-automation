# AP Payment Notification Automation

An Excel-based accounts payable automation that integrates Microsoft Dynamics 365 Business Central data with Power Query and VBA to prepare payment notification emails in Microsoft Outlook.

The project was built to automate a repetitive AP process: identifying electronically paid vendor invoices, matching payments to the invoices they closed, organizing the results by entity and vendor, and generating formatted payment notification emails for review.

## Overview

The workflow combines:

- Microsoft Dynamics 365 Business Central OData web services
- Excel Power Query
- VBA
- Microsoft Outlook
- Multi-entity AP processing

The production implementation was designed for multiple Business Central companies. This public repository has been sanitized and uses generic entity names and fictional demo data.

No production company data, credentials, tenant IDs, vendor information, or email addresses are included.

## Workflow

The automation follows this general process:

1. Retrieve Vendor Ledger Entry data from Business Central.
2. Identify electronic payments posted through bank accounts.
3. Retrieve invoice Vendor Ledger Entries.
4. Match payments to invoices using Vendor No. and Closed at Date.
5. Merge the results with an AP vendor/contact lookup table.
6. Combine results across multiple entities.
7. Group invoices by entity and vendor.
8. Generate formatted HTML payment notification emails in Outlook.
9. Display the emails for user review before sending.

## Architecture

```text
Dynamics 365 Business Central
          |
          v
     OData Web Service
          |
          v
      Power Query
          |
          +--> Payment Entries
          |
          +--> Invoice Entries
          |
          v
 Payment / Invoice Matching
          |
          v
     Vendor Lookup
          |
          v
 Multi-Entity Consolidation
          |
          v
        Excel
          |
          v
         VBA
          |
          v
 Microsoft Outlook
          |
          v
 Payment Notification Email
```

## Multi-Entity Design

The public version uses three generic entities:

- ENT1
- ENT2
- ENT3

Each entity has its own Business Central company parameter and Power Query processing path.

The entity results are combined into a single final AP email dataset before being passed to VBA.

## Business Central Configuration

Business Central connection information is parameterized rather than hard-coded.

The Power Query implementation uses the following parameters:

```text
TenantID
Environment
ENT1_Company
ENT2_Company
ENT3_Company
```

Example placeholders:

```text
TenantID = YOUR-TENANT-ID
Environment = YOUR-ENVIRONMENT
ENT1_Company = YOUR-ENTITY-1-COMPANY
```

This allows the same query architecture to be configured for different Business Central environments without embedding organization-specific connection information in the source code.

## Power Query Logic

The Power Query layer performs the primary data preparation.

### Payment Entries

Electronic payment entries are identified from Vendor Ledger Entries using criteria including:

- Payment Method Code is not CHECK
- Balancing Account Type is Bank Account

### Invoice Entries

Invoice Vendor Ledger Entries are filtered using:

```text
Document Type = Invoice
```

### Payment-to-Invoice Matching

Payments are matched to invoices using:

```text
Vendor_No
+
Closed_at_Date
```

This allows invoices closed by a payment to be associated with the corresponding vendor payment.

### Vendor Lookup

The matched results are merged with an Excel-maintained vendor/contact table containing:

- Entity
- Vendor
- Primary email
- Additional email recipients
- Greeting
- Contact name

### Final Dataset

Results from all entities are combined into `AP_Email_Final`, which becomes the input for the VBA email automation.

## VBA Email Automation

The VBA module reads the final AP dataset and:

- Groups records by entity
- Groups invoices by vendor within each entity
- Builds HTML invoice tables
- Calculates vendor totals
- Builds To and CC recipient lists
- Creates one Outlook email per entity
- Inserts the AP notification above the user's Outlook signature

For safety, the public version defaults to:

```vb
Private Const SEND_EMAILS As Boolean = False
```

This causes emails to be displayed for review rather than automatically sent.

## Demo Data

The workbook includes fictional demo Vendor Ledger Entry data for:

- ENT1
- ENT2
- ENT3

The demo records illustrate the relationship between a payment and the invoices associated with that payment:

```text
Payment
   |
   +--> Invoice 1
   +--> Invoice 2
```

All demo vendors, invoice numbers, payment numbers, bank accounts, amounts, and email addresses are fictional.

Email examples use the reserved `example.com` domain.

## Repository Structure

```text
ap-payment-notification-automation/
|
|-- AP_Payment_Notification_Github.xlsm
|
|-- PowerQuery/
|   |-- ENT1_AP_VLE.m
|   |-- ENT1_AP_Invoices.m
|   |-- ENT1_AP_Email.m
|   |-- ENT2_AP_VLE.m
|   |-- ENT2_AP_Invoices.m
|   |-- ENT2_AP_Email.m
|   |-- ENT3_AP_VLE.m
|   |-- ENT3_AP_Invoices.m
|   |-- ENT3_AP_Email.m
|   |-- AP_Vendors.m
|   `-- AP_Email_Final.m
|
|-- VBA/
|   `-- AP_Payment_Notifications.bas
|
|-- README.md
`-- .gitignore
```

## Technologies

**Microsoft Dynamics 365 Business Central**  
Source ERP and Vendor Ledger Entry data.

**OData**  
Provides the web-service interface between Business Central and Excel.

**Power Query / M**  
Retrieves, filters, joins, transforms, and consolidates AP data.

**Microsoft Excel**  
Provides the user interface, configuration tables, control sheet, and final AP dataset.

**VBA**  
Groups the prepared AP data and builds the notification emails.

**Microsoft Outlook**  
Displays or sends the final formatted vendor payment notifications.

## Security

The public repository has been sanitized.

The repository does not contain:

- Production Business Central tenant IDs
- Production environment names
- Real company/entity names
- Real vendor names
- Real vendor email addresses
- Real employee email addresses
- Production AP transaction data

Business Central connection information must be supplied through the workbook parameters before connecting the workbook to an actual environment.

## Purpose

This project demonstrates how accounting workflow knowledge can be combined with Business Central, Power Query, Excel, VBA, and Outlook to automate a multi-step accounts payable process while maintaining a review step before external communication.