let
    Source = Table.NestedJoin(
        ENT1_AP_VLE,
        {"Vendor_No", "Closed_at_Date"},
        ENT1_AP_Invoices,
        {"Vendor_No", "Closed_at_Date"},
        "ENT1_AP_Invoices",
        JoinKind.LeftOuter
    ),

    #"Expanded ENT1_AP_Invoices" = Table.ExpandTableColumn(
        Source,
        "ENT1_AP_Invoices",
        {"Document_No", "External_Document_No", "Original_Amount"},
        {"Document_No.1", "External_Document_No.1", "Original_Amount.1"}
    ),

    #"Renamed Columns" = Table.RenameColumns(
        #"Expanded ENT1_AP_Invoices",
        {
            {"Document_No.1", "Internal_Invoice_No"},
            {"External_Document_No.1", "Invoice_No"},
            {"Original_Amount.1", "Invoice_Amount"}
        }
    ),

    #"Added Entity" = Table.AddColumn(
        #"Renamed Columns",
        "Entity",
        each "ENT1"
    ),

    #"Merged Queries" = Table.NestedJoin(
        #"Added Entity",
        {"Entity", "Vendor_Name"},
        AP_Vendors,
        {"Company", "Vendor"},
        "AP_Vendors",
        JoinKind.Inner
    ),

    #"Expanded AP_Vendors" = Table.ExpandTableColumn(
        #"Merged Queries",
        "AP_Vendors",
        {"Email 1", "Email 2", "Email 3", "Dear", "Name"},
        {"Email 1", "Email 2", "Email 3", "Dear", "Name"}
    ),

    #"Calculated Absolute Value" = Table.TransformColumns(
        #"Expanded AP_Vendors",
        {{"Invoice_Amount", Number.Abs, type number}}
    )
in
    #"Calculated Absolute Value"