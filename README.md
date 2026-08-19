\# AP Payment Notification Automation



An Excel-based accounts payable automation that integrates Microsoft Dynamics 365 Business Central data with Power Query and VBA to prepare payment notification emails in Microsoft Outlook.



The project was built to automate a repetitive AP process: identifying electronically paid vendor invoices, matching payments to the invoices they closed, organizing the results by entity and vendor, and generating formatted payment notification emails for review.



\## Overview



The workflow combines:



\- Microsoft Dynamics 365 Business Central OData web services

\- Excel Power Query

\- VBA

\- Microsoft Outlook

\- Multi-entity AP processing



The production implementation was designed for multiple Business Central companies. This public repository has been sanitized and uses generic entity names and fictional demo data.



No production company data, credentials, tenant IDs, vendor information, or email addresses are included.



\## Workflow



The automation follows this general process:



1\. Retrieve Vendor Ledger Entry data from Business Central.

2\. Identify electronic payments posted through bank accounts.

3\. Retrieve invoice Vendor Ledger Entries.

4\. Match payments to invoices using Vendor No. and Closed at Date.

5\. Merge the results with an AP vendor/contact lookup table.

6\. Combine results across multiple entities.

7\. Group invoices by entity and vendor.

8\. Generate formatted HTML payment notification emails in Outlook.

9\. Display the emails for user review before sending.



\## Architecture



```text

Dynamics 365 Business Central

&#x20;         |

&#x20;         v

&#x20;    OData Web Service

&#x20;         |

&#x20;         v

&#x20;     Power Query

&#x20;         |

&#x20;         +--> Payment Entries

&#x20;         |

&#x20;         +--> Invoice Entries

&#x20;         |

&#x20;         v

&#x20;Payment / Invoice Matching

&#x20;         |

&#x20;         v

&#x20;    Vendor Lookup

&#x20;         |

&#x20;         v

&#x20;Multi-Entity Consolidation

&#x20;         |

&#x20;         v

&#x20;       Excel

&#x20;         |

&#x20;         v

&#x20;        VBA

&#x20;         |

&#x20;         v

&#x20;Microsoft Outlook

&#x20;         |

&#x20;         v

&#x20;Payment Notification Email





**Multi-Entity Design**



The public version uses three generic entities:



ENT1

ENT2

ENT3



Each entity has its own Business Central company parameter and Power Query processing path.



The entity results are combined into a single final AP email dataset before being passed to VBA.



**Business Central Configuration**



Business Central connection information is parameterized rather than hard-coded.



The Power Query implementation uses the following parameters:



TenantID

Environment

ENT1\_Company

ENT2\_Company

ENT3\_Company





Example placeholders:

TenantID = YOUR-TENANT-ID

Environment = YOUR-ENVIRONMENT

ENT1\_Company = YOUR-ENTITY-1-COMPANY



This allows the same query architecture to be configured for different Business Central environments without embedding organization-specific connection information in the source code.



**Power Query Logic**



The Power Query layer performs the primary data preparation.



**Payment Entries**



Electronic payment entries are identified from Vendor Ledger Entries using criteria including:



Payment Method Code is not CHECK

Balancing Account Type is Bank Account



**Invoice Entries**



Invoice Vendor Ledger Entries are filtered using:

Document Type = Invoice



**Payment-to-Invoice Matching**



Payments are matched to invoices using:

Vendor\_No

\+

Closed\_at\_Date



This allows invoices closed by a payment to be associated with the corresponding vendor payment.



**Vendor Lookup**



The matched results are merged with an Excel-maintained vendor/contact table containing:



Entity

Vendor

Primary email

Additional email recipients

Greeting

Contact name



**Final Dataset**



Results from all entities are combined into AP\_Email\_Final, which becomes the input for the VBA email automation.



**VBA Email Automation**



The VBA module reads the final AP dataset and:



Groups records by entity

Groups invoices by vendor within each entity

Builds HTML invoice tables

Calculates vendor totals

Builds To and CC recipient lists

Creates one Outlook email per entity

Inserts the AP notification above the user's Outlook signature



For safety, the public version defaults to:

Private Const SEND\_EMAILS As Boolean = False



This causes emails to be displayed for review rather than automatically sent.



**Demo Data**



The workbook includes fictional demo Vendor Ledger Entry data for:

ENT1

ENT2

ENT3



The demo records illustrate the relationship between:

Payment

&#x20;  |

&#x20;  +--> Invoice 1

&#x20;  +--> Invoice 2



All demo vendors, invoice numbers, payment numbers, bank accounts, amounts, and email addresses are fictional.



Email examples use the reserved example.com domain.



**Repository Structure**



ap-payment-notification-automation/

|

|-- AP\_Payment\_Notification\_Github.xlsm

|

|-- PowerQuery/

|   |-- ENT1\_AP\_VLE.m

|   |-- ENT1\_AP\_Invoices.m

|   |-- ENT1\_AP\_Email.m

|   |-- ENT2\_AP\_VLE.m

|   |-- ENT2\_AP\_Invoices.m

|   |-- ENT2\_AP\_Email.m

|   |-- ENT3\_AP\_VLE.m

|   |-- ENT3\_AP\_Invoices.m

|   |-- ENT3\_AP\_Email.m

|   |-- AP\_Vendors.m

|   `-- AP\_Email\_Final.m

|

|-- VBA/

|   `-- AP\_Payment\_Notifications.bas

|

|-- README.md

`-- .gitignore





**Technologies**



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



**Security**



The public repository has been sanitized.



The repository does not contain:



Production Business Central tenant IDs

Production environment names

Real company/entity names

Real vendor names

Real vendor email addresses

Real employee email addresses

Production AP transaction data



Business Central connection information must be supplied through the workbook parameters before connecting the workbook to an actual environment.



**Purpose**



This project demonstrates how accounting workflow knowledge can be combined with Business Central, Power Query, Excel, VBA, and Outlook to automate a multi-step accounts payable process while maintaining a review step before external communication.

