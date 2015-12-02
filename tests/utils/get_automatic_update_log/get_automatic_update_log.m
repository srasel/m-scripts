/*
	Script Name	: set automatic/frequent updating of datasets & srcs
	Author		: Md. Shahnewaz Rasel
	Input	: Folder Path		  
	Output	: Track all files in that folder
	GitHub	: #06
*/
(Folder_Path as text)=>
let
    Source = Folder.Files(Folder_Path),
    #"Added Custom" = Table.AddColumn(Source, "File Path", each [Folder Path] & [Name]),
    #"Reordered Columns" = Table.ReorderColumns(#"Added Custom",{"Content", "Name", "Extension", "File Path", "Date accessed", "Date modified", "Date created", "Attributes", "Folder Path"}),
    #"Removed Other Columns" = Table.SelectColumns(#"Reordered Columns",{"File Path", "Date accessed", "Date modified", "Date created"}),
    Output = Table.Sort(#"Removed Other Columns",{{"Date modified", Order.Descending}})
in
    Output