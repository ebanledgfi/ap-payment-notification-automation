let
    Source = Excel.CurrentWorkbook(){[Name="tblAPVendors"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"Company", type text}, {"Vendor", type text}, {"Email 1", type text}, {"Email 2", type text}, {"Email 3", type any}, {"Dear", type text}, {"Name", type text}})
in
    #"Changed Type"