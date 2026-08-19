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