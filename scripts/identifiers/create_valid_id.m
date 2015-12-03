/*
    Script Name    : Clean latin characters to make it readable.
    Author        : Md. Shahnewaz Rasel
    Input    : will be a 3 columns spreadsheet with a Table name defined as "Source"
              ideally with one single column. Script will prompt for the name 
    Output    : Clean readable words.
    GitHub    : #21
*/
(Latin_Column as text)=>
let
    // Assumption is: Source will be a excel file
    // with a Sheet name: Source
    Source = Excel.CurrentWorkbook(){[Name="Source"]}[Content],
    DataTable = Source, 
	AccCharsList = "öŠŽšžŸÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÐÑÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïðñòóôõöùúûüýÿ",
	RegCharsList = "oSZszYAAAAAACEEEEIIIIDNOOOOOUUUUYaaaaaaceeeeiiiidnooooouuuuyy",
	SpecialCharsList =", |*()[]",
	
	input = Table.RenameColumns(DataTable,{{Latin_Column, "Geocol"}}),
	
	replaceFunc = (inText as text, indx as number) as text =>
	let	
		AccChar= Text.Range(AccCharsList,indx,1),
		RegChar= Text.Range(RegCharsList,indx,1),
		SpecialChar = (
		        if indx < Text.Length(SpecialCharsList) then 
		            Text.Range(SpecialCharsList,indx,1)
				else
				    ""
		),
		latinReplaced = Text.Replace(inText,AccChar,RegChar),
		finalString = Text.Replace(latinReplaced,SpecialChar,"_")
	in
		finalString,
	CleanLatinChar = [
 		Loop = 
			(inText as text, indx as number) as text=>
			if indx >= Text.Length(AccCharsList) then
				inText
			else 
				@Loop(replaceFunc(inText,indx),indx+1)
	],
	FinalTab = Table.TransformRows(input, each CleanLatinChar[Loop]([Geocol],0)),
    #"Converted to Table" = Table.FromList(FinalTab, Splitter.SplitByNothing(), null, null, ExtraValues.Error),
    #"Lowercased Text" = Table.TransformColumns(#"Converted to Table",{},Text.Lower),
    ReplaceDoubleUnderscore = Table.ReplaceValue(#"Lowercased Text","__","_",Replacer.ReplaceText,{"Column1"}),
    Output = Table.RenameColumns(ReplaceDoubleUnderscore,{{"Column1", Latin_Column}})
in
    Output