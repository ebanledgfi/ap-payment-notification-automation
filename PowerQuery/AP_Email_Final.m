let
    Source = Table.Combine({ENT1_AP_Email, ENT2_AP_Email, ENT3_AP_Email}),
    #"Removed Other Columns" = Table.SelectColumns(
        Source,
        {
            "Posting_Date",
            "Document_No",
            "External_Document_No",
            "Vendor_No",
            "Vendor_Name",
            "Payment_Method_Code",
            "Original_Amount",
            "Invoice_No",
            "Invoice_Amount",
            "Entity",
            "Email 1",
            "Email 2",
            "Email 3",
            "Dear",
            "Name"
        }
    ),
    #"Reordered Columns" = Table.ReorderColumns(
        #"Removed Other Columns",
        {
            "Entity",
            "Posting_Date",
            "Document_No",
            "Vendor_No",
            "Vendor_Name",
            "Payment_Method_Code",
            "Original_Amount",
            "Invoice_No",
            "Invoice_Amount",
            "External_Document_No",
            "Email 1",
            "Email 2",
            "Email 3",
            "Dear",
            "Name"
        }
    ),
    #"Removed Columns" = Table.RemoveColumns(
        #"Reordered Columns",
        {"External_Document_No"}
    ),
    #"Renamed Columns" = Table.RenameColumns(
        #"Removed Columns",
        {
            {"Posting_Date", "Payment Date"},
            {"Vendor_No", "Vendor No"},
            {"Vendor_Name", "Vendor"},
            {"Payment_Method_Code", "Payment Method"},
            {"Original_Amount", "Payment Amount"},
            {"Invoice_No", "Invoice No"},
            {"Invoice_Amount", "Invoice Amount"}
        }
    )
in
    #"Renamed Columns"