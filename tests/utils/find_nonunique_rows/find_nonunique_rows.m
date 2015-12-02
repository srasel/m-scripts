let
    Source = Excel.CurrentWorkbook(){[Name="Source"]}[Content],
    #"Changed Type" = Table.TransformColumnTypes(Source,{{"geo", type text}, {"year", Int64.Type}, {"pop", Int64.Type}}),
	cols = Table.ColumnNames(#"Changed Type"),
	#"Grouped Rows" = Table.Group(#"Changed Type", {"geo", "year", "pop"}, {{"allRows", each _, type table}}),
	
	RankFunction = (tabletorank as table) as table =>
     let
      SortRows = Table.Sort(tabletorank,{{"geo", Order.Ascending}, {"year", Order.Ascending}, {"pop", Order.Ascending}}),
      AddIndex = Table.AddIndexColumn(SortRows, "Rank", 1, 1)
     in
      AddIndex,
    //Apply that function to the AllRows column
    AddedRank = Table.TransformColumns(#"Grouped Rows", {"allRows", each RankFunction(_)}),
    #"Removed Other Columns" = Table.SelectColumns(AddedRank,{"allRows"}),
    #"Expanded allRows" = Table.ExpandTableColumn(#"Removed Other Columns", "allRows", {"geo", "year", "pop", "Rank"}, {"geo", "year", "pop", "Rank"}),
    #"Filtered Rows" = Table.SelectRows(#"Expanded allRows", each ([Rank] = 2)),
    #"Removed Columns" = Table.RemoveColumns(#"Filtered Rows",{"Rank"}),
	renames = List.Transform(
        Table.ColumnNames(#"Removed Columns"),
        each {_, Text.Combine(_,",")})
in
    renames