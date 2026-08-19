let
    CompanyEncoded = Uri.EscapeDataString(ENT2_Company),
    Source = OData.Feed(
        "https://api.businesscentral.dynamics.com/v2.0/"
        & TenantID
        & "/"
        & Environment
        & "/ODataV4/Company('"
        & CompanyEncoded
        & "')/Vendor_Ledger_Entries_Excel",
        null,
        [Implementation="2.0"]
    ),
    #"Removed Other Columns" = Table.SelectColumns(Source,{"Entry_No", "Posting_Date", "Document_Date", "Document_Type", "Document_No", "External_Document_No", "Vendor_No", "Vendor_Name", "Payment_Method_Code", "Original_Amount", "Original_Amt_LCY", "Bal_Account_Type", "Bal_Account_No", "Open", "Reversed", "Closed_at_Date"}),
    #"Filtered Rows" = Table.SelectRows(#"Removed Other Columns", each ([Payment_Method_Code] <> "CHECK") and ([Bal_Account_Type] = "Bank Account"))
in
    #"Filtered Rows"