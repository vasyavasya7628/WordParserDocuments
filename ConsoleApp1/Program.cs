namespace ConsoleApp1;

class Program
{
    static void Main()
    {
        string wordFilePath = "new.docx";
        string htmlOutput = ConsoleApp1.ConvertToHtml(wordFilePath);

        Console.WriteLine(htmlOutput);
    }
}