using System.Text.RegularExpressions;
using System.Web;
using Xceed.Words.NET;

namespace ConsoleApp1;

public class ConsoleApp1
{
    public static string ConvertToHtml(string filePath)
    {
        string styleHead = Class1.Head();
        using (DocX doc = DocX.Load(filePath))
        {
            string html = styleHead;
            string[] patterns = new string[]
            {
                @"_{2,}", // Паттерн для поиска подчеркиваний (более 2 символов)
                @"^(?!_{2,}).+$" // Паттерн для поиска строки, не содержащей подчеркивание
            };
            
            foreach (var paragraph in doc.Paragraphs)
            {
                string paragraphText = HtmlEncode(paragraph.Text);
                string formattedHtml = " ";
                if (paragraphText.Length > 10)
                {

                    string[] lines = Regex.Split(paragraphText, @"\r\n");
                    foreach (string line in lines) {
                        Console.WriteLine(line);
                    }
                }
                

                //formattedHtml = formatHtml(patterns, paragraphText, formattedHtml, ref shouldAddBlock);
                
    
                //html += paragraphText;
            }
            

            html += "</body></html>";
            return html;
        }
    }

    private static string formatHtml(string[] patterns, string paragraphText, string formattedHtml)
    {
        foreach (var pattern in patterns)
        {
            try
            {
                Match match = Regex.Match(paragraphText, pattern);

                switch (pattern)
                {
                    case @"_{2,}":
                        string replace = match.Value.Replace("_", ""); // Удаление подчеркиваний
                        formattedHtml +=
                            $"<table width=\"100%\">\r\n" +
                            $"  <tr>\r\n" +
                            $"    <td width=\"60%\">\r\n" +
                            $"    </td>\r\n" +
                            $"    <td class=\"under_line\" width=\"40%\">\r\n" +
                            $"      <span style=\"border-bottom: 3px solid white;\">{replace}</span><br></td>\r\n" +
                            $"  </tr>\r\n" +
                            $"</table>";
                        break;

                    case @"^(?!_{2,}).+$":
                        formattedHtml +=
                            $"<table width=\"100%\">\r\n" +
                            $"  <tr>\r\n" +
                            $"    <td width=\"100%\"\">{paragraphText}</td>\r\n" +
                            $"  </tr>\r\n" +
                            $"</table>";

                        break;
                }
            }
            catch (RegexParseException ex)
            {
                Console.WriteLine("Ошибка в регулярном выражении: " + ex.Message);
            }
        }

        return formattedHtml;
    }

    private static string HtmlEncode(string input)
    {
        return HttpUtility.HtmlEncode(input);
    }
}