/*
	Script Name	: Interpolate / Extrapolate to find the missing values
	Author		: Md. Shahnewaz Rasel
	Input	: will be a 3 columns spreadsheet with a Table name defined as "Source"
			  ideally columns will be geo, year, indicator_name or geo, age, indicator_name			  
	Output	: Script output with the necessary columns
	GitHub	: #08
*/
(Ordinal_Column as text, Interpolation_Type as text)=>
let
	// Assumption is: Source will be a excel file
	// with a Sheet with Table name: Source
	Source = Excel.CurrentWorkbook(){[Name="Source"]}[Content],
    DataTable = Source, 
	// Change the year/age/etc. ordinal column name to common name "ordinalColumn"
	// to use further.
	RenameOrdinalColumn = Table.RenameColumns(DataTable,{{Ordinal_Column, "ordinalColumn"}}),
    CollectDimensionColumns = {"geo","ordinalColumn"}, 
    ListOfAllColumns = Table.ColumnNames(RenameOrdinalColumn),
    IndicatorColumns = List.RemoveMatchingItems(ListOfAllColumns,CollectDimensionColumns),	
    UnpivotedTable = Table.Unpivot(RenameOrdinalColumn,IndicatorColumns,"Attribute","Value"),
	
	// Add _interpol_ordinalcolumn_interpolationtype property to all indicator name
    AddAttributeDummyCol = Table.AddColumn(UnpivotedTable, "Attribute_1", each [Attribute] & "_interpol_" & Ordinal_Column & "_" & Interpolation_Type),
    RemoveExistingAttributeCol = Table.RemoveColumns(AddAttributeDummyCol,{"Attribute"}),
    RenameBackToAttribute = Table.RenameColumns(RemoveExistingAttributeCol,{{"Attribute_1", "Attribute"}}),

    #"Changed Type" = Table.TransformColumnTypes(RenameBackToAttribute,{{"ordinalColumn", Int64.Type}, {"Value", type number}}),
    MaxordinalColumn = Table.AddColumn(#"Changed Type", "MaxordinalColumn", each List.Max(Table.Column(#"Changed Type","ordinalColumn"))),
    ReplicateRows = [
 		Factorial = (singleRow as list) as table=>
                if singleRow{2} = singleRow{4} then 
			Table.TransformColumnTypes(
					Table.FromRecords({[geo=singleRow{0}, Attribute=singleRow{1}, ordinalColumn=singleRow{2}, Value=singleRow{3}, MaxordinalColumn=singleRow{4}]})
			,{{"geo",type text},{"Attribute", type text},{"ordinalColumn", Int64.Type},{"Value", type number},{"MaxordinalColumn", Int64.Type}})
		else 
			Table.Combine(
				{Table.TransformColumnTypes(
					Table.FromRecords({[geo=singleRow{0}, Attribute=singleRow{1}, ordinalColumn=singleRow{2}, Value=singleRow{3}, MaxordinalColumn=singleRow{4}]})
					,{{"geo",type text},{"Attribute", type text},{"ordinalColumn", Int64.Type},{"Value", type number},{"MaxordinalColumn", Int64.Type}})
				,@Factorial({singleRow{0},singleRow{1},singleRow{2}+1,null,singleRow{4}})
				})
		
                

	],
	FinalTab = Table.TransformRows(MaxordinalColumn, each ReplicateRows[Factorial]({[geo],[Attribute],[ordinalColumn],[Value],[MaxordinalColumn]})),
	A = Table.Combine(FinalTab),
	B = Table.Combine({MaxordinalColumn}),
	C = Table.Combine({A,B}),
    #"Grouped Rows" = Table.Group(C, {"geo", "Attribute", "ordinalColumn"}, {{"value", each List.Max([Value]), type number}}),
    startordinalColumn = Table.AddColumn(#"Grouped Rows", "startordinalColumn", each if [value] = null then null else [ordinalColumn]),
    endordinalColumn = Table.AddColumn(startordinalColumn, "endordinalColumn", each if [value] = null then null else [ordinalColumn]),
    startVal = Table.AddColumn(endordinalColumn, "startVal", each [value]),
    endVal = Table.AddColumn(startVal, "endVal", each [value]),
    GrpVal = Table.Group(endVal, {"geo", "Attribute"}, {{"AllRows", each _, type table}}),

    FillValues = (tabletorank as table) as table =>
    	let
      		FillUpordinalColumn = Table.FillUp(tabletorank,{"endordinalColumn"}),
      		FillDownordinalColumn = Table.FillDown(FillUpordinalColumn,{"startordinalColumn"}),
      		FillUpVal = Table.FillUp(FillDownordinalColumn,{"endVal"}),
      		FillDownVal = Table.FillDown(FillUpVal,{"startVal"})
     	in
    FillDownVal,
    AllTab = Table.TransformColumns(GrpVal, {"AllRows", each FillValues(_)}),
    #"Expanded AllRows" = Table.ExpandTableColumn(AllTab, "AllRows", {"ordinalColumn", "value", "startordinalColumn", "endordinalColumn", "startVal", "endVal"}, {"AllRows.ordinalColumn", "AllRows.value", "AllRows.startordinalColumn", "AllRows.endordinalColumn", "AllRows.startVal", "AllRows.endVal"}),
    AddNotesCol = Table.AddColumn(#"Expanded AllRows", "notes", each if [AllRows.value] = null then 
		if [AllRows.endordinalColumn] <> null 
			then "Value is interpolated linearly between " & Text.From([AllRows.startordinalColumn]) & " and " & Text.From([AllRows.endordinalColumn]) & ", from " & Text.From([AllRows.startVal]) & " to " & Text.From([AllRows.endVal]) else "Value is extrapolated linearly from " & Text.From([AllRows.startordinalColumn]) & " with value " & Text.From([AllRows.startVal])
		else ""
	),    
    AddRankCol = Table.AddColumn(AddNotesCol, "rank", each [AllRows.ordinalColumn] - [AllRows.startordinalColumn]),
	AddDoubtCol = Table.AddColumn(AddRankCol, "doubt", each if [rank] = 0 then "" else Text.From([rank] * 5) & "%"),

    AddStepCol = Table.AddColumn(AddDoubtCol, "step", each [AllRows.endordinalColumn] - [AllRows.startordinalColumn]),
    CalculateInterimValue = Table.AddColumn(AddStepCol, "interimVal", each if [AllRows.value] = null then ([AllRows.startVal] + ([AllRows.endVal]-[AllRows.startVal])*[rank]/[step]) else [AllRows.value]),
    FinalValue = Table.AddColumn(CalculateInterimValue, "finalValue", each if [AllRows.endordinalColumn] = null then [AllRows.startVal] else [interimVal]),
    RenameValueCol = Table.RenameColumns(FinalValue,{{"finalValue", "value"}}),
    RemoveOtherCols = Table.RemoveColumns(RenameValueCol,{"AllRows.value", "AllRows.startordinalColumn", "AllRows.endordinalColumn", "AllRows.startVal", "AllRows.endVal", "rank", "step", "interimVal"}),
    #"Reordered Columns" = Table.ReorderColumns(RemoveOtherCols,{"geo", "AllRows.ordinalColumn", "Attribute", "value","notes","doubt"}),
    #"Pivoted Column" = Table.Pivot(#"Reordered Columns", List.Distinct(#"Reordered Columns"[Attribute]), "Attribute", "value", List.Sum),
    #"Renamed Columns_2" = Table.RenameColumns(#"Pivoted Column",{{"notes", "interpol_" & Ordinal_Column & "_" & Interpolation_Type & ".notes"}, {"doubt", "interpol_" & Ordinal_Column & "_" & Interpolation_Type & ".doubt"}}),
    Output = Table.RenameColumns(#"Renamed Columns_2",{{"AllRows.ordinalColumn", Ordinal_Column}})
in
    Output