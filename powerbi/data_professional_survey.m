// Data Professional Survey: Power Query (M) cleaning steps
//
// How to use: in Power BI Desktop open Transform data > Advanced Editor, paste this script,
// and change the file path in the Source step to where you saved data/Power BI - Final Project.xlsx
//
// What the steps do:
//   - Promote headers and set data types
//   - Keep only the text before "(" or ":" in the job title, industry, country, new-job
//     priority and programming-language answers, so "Other (Please Specify):..." becomes "Other"
//   - Turn the salary bracket text (e.g. "41k-65k") into a low and high number, then compute
//     "Average Salary" as the bracket midpoint in $k: (low + high) / 2. The top bracket "225k+" is counted as 225.
//   - Replace blank education values with "Other"
//   - Remove the empty columns: Browser, OS, City, Country, Referrer
//
let
    Source = Excel.Workbook(File.Contents("C:\path\to\Power BI - Final Project.xlsx"), null, true),
    #"Data Professional Survey_Sheet" = Source{[Item="Data Professional Survey",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(#"Data Professional Survey_Sheet", [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"Unique ID", type text}, {"Email", type text}, {"Date Taken (America/New_York)", type date}, {"Time Taken (America/New_York)", type time}, {"Browser", type text}, {"OS", type text}, {"City", type text}, {"Country", type text}, {"Referrer", type text}, {"Time Spent", type time}, {"Q1 - Which Title Best Fits your Current Role?", type text}, {"Q2 - Did you switch careers into Data?", type text}, {"Q3 - Current Yearly Salary (in USD)", type text}, {"Q4 - What Industry do you work in?", type text}, {"Q5 - Favorite Programming Language", type text}, {"Q6 - How Happy are you in your Current Position with the following? (Salary)", Int64.Type}, {"Q6 - How Happy are you in your Current Position with the following? (Work/Life Balance)", Int64.Type}, {"Q6 - How Happy are you in your Current Position with the following? (Coworkers)", Int64.Type}, {"Q6 - How Happy are you in your Current Position with the following? (Management)", Int64.Type}, {"Q6 - How Happy are you in your Current Position with the following? (Upward Mobility)", Int64.Type}, {"Q6 - How Happy are you in your Current Position with the following? (Learning New Things)", Int64.Type}, {"Q7 - How difficult was it for you to break into Data?", type text}, {"Q8 - If you were to look for a new job today, what would be the most important thing to you?", type text}, {"Q9 - Male/Female?", type text}, {"Q10 - Current Age", Int64.Type}, {"Q11 - Which Country do you live in?", type text}, {"Q12 - Highest Level of Education", type text}, {"Q13 - Ethnicity", type text}}),
    #"Split Column by Delimiter" = Table.SplitColumn(#"Changed Type", "Q1 - Which Title Best Fits your Current Role?", Splitter.SplitTextByEachDelimiter({"("}, QuoteStyle.Csv, false), {"Q1 - Which Title Best Fits your Current Role?.1", "Q1 - Which Title Best Fits your Current Role?.2"}),
    #"Changed Type1" = Table.TransformColumnTypes(#"Split Column by Delimiter",{{"Q1 - Which Title Best Fits your Current Role?.1", type text}, {"Q1 - Which Title Best Fits your Current Role?.2", type text}}),
    #"Removed Columns" = Table.RemoveColumns(#"Changed Type1",{"Q1 - Which Title Best Fits your Current Role?.2"}),
    #"Split Column by Delimiter1" = Table.SplitColumn(#"Removed Columns", "Q5 - Favorite Programming Language", Splitter.SplitTextByEachDelimiter({":"}, QuoteStyle.Csv, false), {"Q5 - Favorite Programming Language.1", "Q5 - Favorite Programming Language.2"}),
    #"Changed Type2" = Table.TransformColumnTypes(#"Split Column by Delimiter1",{{"Q5 - Favorite Programming Language.1", type text}, {"Q5 - Favorite Programming Language.2", type text}}),
    #"Removed Columns1" = Table.RemoveColumns(#"Changed Type2",{"Q5 - Favorite Programming Language.2"}),
    #"Duplicated Column" = Table.DuplicateColumn(#"Removed Columns1", "Q3 - Current Yearly Salary (in USD)", "Q3 - Current Yearly Salary (in USD) - Copy"),
    #"Split Column by Character Transition" = Table.SplitColumn(#"Duplicated Column", "Q3 - Current Yearly Salary (in USD) - Copy", Splitter.SplitTextByCharacterTransition({"0".."9"}, (c) => not List.Contains({"0".."9"}, c)), {"Q3 - Current Yearly Salary (in USD) - Copy.1", "Q3 - Current Yearly Salary (in USD) - Copy.2", "Q3 - Current Yearly Salary (in USD) - Copy.3"}),
    #"Removed Columns2" = Table.RemoveColumns(#"Split Column by Character Transition",{"Q3 - Current Yearly Salary (in USD) - Copy.3"}),
    #"Replaced Value" = Table.ReplaceValue(#"Removed Columns2","k","",Replacer.ReplaceText,{"Q3 - Current Yearly Salary (in USD) - Copy.2"}),
    #"Replaced Value1" = Table.ReplaceValue(#"Replaced Value","-","",Replacer.ReplaceText,{"Q3 - Current Yearly Salary (in USD) - Copy.2"}),
    #"Replaced Value2" = Table.ReplaceValue(#"Replaced Value1","+","225",Replacer.ReplaceText,{"Q3 - Current Yearly Salary (in USD) - Copy.2"}),
    #"Added Custom" = Table.AddColumn(#"Replaced Value2", "Average Salary", each ([#"Q3 - Current Yearly Salary (in USD) - Copy.1"]) + ([#"Q3 - Current Yearly Salary (in USD) - Copy.2"])),
    #"Changed Type3" = Table.TransformColumnTypes(#"Added Custom",{{"Q3 - Current Yearly Salary (in USD) - Copy.2", Int64.Type}, {"Q3 - Current Yearly Salary (in USD) - Copy.1", Int64.Type}, {"Average Salary", Int64.Type}}),
    #"Removed Columns3" = Table.RemoveColumns(#"Changed Type3",{"Average Salary"}),
    #"Added Custom1" = Table.AddColumn(#"Removed Columns3", "Average Salary", each ([#"Q3 - Current Yearly Salary (in USD) - Copy.1"]) + ([#"Q3 - Current Yearly Salary (in USD) - Copy.2"])),
    #"Removed Columns4" = Table.RemoveColumns(#"Added Custom1",{"Average Salary"}),
    #"Added Custom2" = Table.AddColumn(#"Removed Columns4", "Average Salary", each ([#"Q3 - Current Yearly Salary (in USD) - Copy.1"] + [#"Q3 - Current Yearly Salary (in USD) - Copy.2"]) / 2),
    #"Removed Columns5" = Table.RemoveColumns(#"Added Custom2",{"Q3 - Current Yearly Salary (in USD) - Copy.1", "Q3 - Current Yearly Salary (in USD) - Copy.2"}),
    #"Reordered Columns" = Table.ReorderColumns(#"Removed Columns5",{"Unique ID", "Email", "Date Taken (America/New_York)", "Time Taken (America/New_York)", "Browser", "OS", "City", "Country", "Referrer", "Time Spent", "Q1 - Which Title Best Fits your Current Role?.1", "Q2 - Did you switch careers into Data?", "Q3 - Current Yearly Salary (in USD)", "Average Salary", "Q4 - What Industry do you work in?", "Q5 - Favorite Programming Language.1", "Q6 - How Happy are you in your Current Position with the following? (Salary)", "Q6 - How Happy are you in your Current Position with the following? (Work/Life Balance)", "Q6 - How Happy are you in your Current Position with the following? (Coworkers)", "Q6 - How Happy are you in your Current Position with the following? (Management)", "Q6 - How Happy are you in your Current Position with the following? (Upward Mobility)", "Q6 - How Happy are you in your Current Position with the following? (Learning New Things)", "Q7 - How difficult was it for you to break into Data?", "Q8 - If you were to look for a new job today, what would be the most important thing to you?", "Q9 - Male/Female?", "Q10 - Current Age", "Q11 - Which Country do you live in?", "Q12 - Highest Level of Education", "Q13 - Ethnicity"}),
    #"Split Column by Delimiter2" = Table.SplitColumn(#"Reordered Columns", "Q11 - Which Country do you live in?", Splitter.SplitTextByEachDelimiter({"("}, QuoteStyle.Csv, false), {"Q11 - Which Country do you live in?.1", "Q11 - Which Country do you live in?.2"}),
    #"Changed Type4" = Table.TransformColumnTypes(#"Split Column by Delimiter2",{{"Q11 - Which Country do you live in?.1", type text}, {"Q11 - Which Country do you live in?.2", type text}}),
    #"Filtered Rows" = Table.SelectRows(#"Changed Type4", each true),
    #"Removed Columns6" = Table.RemoveColumns(#"Filtered Rows",{"Q11 - Which Country do you live in?.2"}),
    #"Split Column by Delimiter3" = Table.SplitColumn(#"Removed Columns6", "Q4 - What Industry do you work in?", Splitter.SplitTextByEachDelimiter({"("}, QuoteStyle.Csv, false), {"Q4 - What Industry do you work in?.1", "Q4 - What Industry do you work in?.2"}),
    #"Changed Type5" = Table.TransformColumnTypes(#"Split Column by Delimiter3",{{"Q4 - What Industry do you work in?.1", type text}, {"Q4 - What Industry do you work in?.2", type text}}),
    #"Removed Columns7" = Table.RemoveColumns(#"Changed Type5",{"Q4 - What Industry do you work in?.2"}),
    #"Split Column by Delimiter4" = Table.SplitColumn(#"Removed Columns7", "Q8 - If you were to look for a new job today, what would be the most important thing to you?", Splitter.SplitTextByEachDelimiter({"("}, QuoteStyle.Csv, false), {"Q8 - If you were to look for a new job today, what would be the most important t", "Q8 - If you were to look for a new job today, what would be the most important.1"}),
    #"Changed Type6" = Table.TransformColumnTypes(#"Split Column by Delimiter4",{{"Q8 - If you were to look for a new job today, what would be the most important t", type text}, {"Q8 - If you were to look for a new job today, what would be the most important.1", type text}}),
    #"Replaced Value3" = Table.ReplaceValue(#"Changed Type6","","Other",Replacer.ReplaceValue,{"Q12 - Highest Level of Education"}),
    #"Removed Columns8" = Table.RemoveColumns(#"Replaced Value3",{"Browser", "OS", "City", "Country", "Referrer"})
in
    #"Removed Columns8"
