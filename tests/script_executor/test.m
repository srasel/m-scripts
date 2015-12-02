()=>
let
    Source = Excel.CurrentWorkbook(){[Name="Source"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"geo", type text}}),
    Output = Table.Group(#"Changed Type", {"geo"}, {{"Count", each Table.RowCount(_), type number}})
in
    Output