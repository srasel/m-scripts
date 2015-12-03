let
	// Define your script path along with script name
    MScriptPath = "C:\path_to_the_script\script_name.m",
	// Read full content as Binary.
    ScriptContentAsBinary = Text.FromBinary(File.Contents(MScriptPath),TextEncoding.Windows),
	// Execute the built-in function to import the content and
	// make it executable
    EvaluatedExpression = Expression.Evaluate(ScriptContentAsBinary, #shared)
in
	EvaluatedExpression