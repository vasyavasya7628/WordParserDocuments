using System;

class Program
{
    static void Main(string[] args)
    {
        string wordFilePath = "new.docx";
        string htmlOutput = WordToHtmlParser.ConvertToHtml(wordFilePath);

        Console.WriteLine(htmlOutput);
    }
}
