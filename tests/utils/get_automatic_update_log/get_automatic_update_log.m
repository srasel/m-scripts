let
    Source = Folder.Files("C:\Users\Dell\Documents\datasets"),
    #"Added Custom" = Table.AddColumn(Source, "File Path", each [Folder Path] & [Name]),
    #"Reordered Columns" = Table.ReorderColumns(#"Added Custom",{"Content", "Name", "Extension", "File Path", "Date accessed", "Date modified", "Date created", "Attributes", "Folder Path"}),
    #"Removed Other Columns" = Table.SelectColumns(#"Reordered Columns",{"File Path", "Date accessed", "Date modified", "Date created"}),
    Output = Table.Sort(#"Removed Other Columns",{{"Date modified", Order.Descending}})
in
    Output