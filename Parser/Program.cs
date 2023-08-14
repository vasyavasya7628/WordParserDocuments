using Parser;

var wordFilePath = "new.docx";
var htmlOutput = WordToHtmlParser.ConvertToHtml(wordFilePath);

Console.WriteLine(htmlOutput);