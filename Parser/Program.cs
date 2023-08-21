using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

//1) если есть underscore тогда заменить количество символов underscore на переменную и оставшиеся на пробелы
//1) если есть underscore тогда заменить количество символов underscore на   border-bottom: 1px solid black размером с количество underscore
class Program
{
    static void Main(string[] args)
    {
        string currentDirectory = Directory.GetCurrentDirectory();
        string fileSearchPattern = "*.docx";
        string inputPath = "*.docx";
        string outputPath = "output.html";
        try
        {
            string[] files = Directory.GetFiles(currentDirectory, fileSearchPattern);
            foreach (var inputfile in files)
            {
                using (WordprocessingDocument doc = WordprocessingDocument.Open(inputfile, false))
                {
                    // Извлечение содержимого из документа Word
                    var content = ExtractContent(doc.MainDocumentPart.Document.Body);

                    // Создание HTML-кода из содержимого
                    string html = GenerateHtmlFromContent(content);

                    // Сохранение HTML-кода в файл
                    File.WriteAllText(Path.GetFileNameWithoutExtension(inputfile) + "_конвертирован" + ".html", html);
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine("404 File not found");
        }
    }

    // Рекурсивная функция для извлечения содержимого из элементов Open XML
    static string ExtractContent(OpenXmlElement element)
    {
        string content = "";
        foreach (var childElement in element.Elements())
        {
            if (childElement is Text)
            {
                content += childElement.InnerText;
            }
            else if (childElement is Run)
            {
                content += ExtractContent(childElement);
            }
            else if (childElement is Paragraph)
            {
                // Получение выравнивания абзаца
                string alignment = GetParagraphAlignment(childElement as Paragraph);

                // Преобразование выравнивания в HTML-атрибут
                string alignmentAttribute = !string.IsNullOrEmpty(alignment) ? $" align=\"{alignment}\"" : "";
                Regex underscorePattern = new Regex(@"\w+_.+");

                content += $"<p{alignmentAttribute}>" + ExtractContent(childElement) + "</p>";
            }
            else if (childElement is Table)
            {
                content += "<table class=\"table_bordered\">" + ExtractTableContent(childElement as Table) + "</table>";
            }
            // Добавьте обработку других типов элементов по аналогии
        }

        return content;
    }

    // Извлечение содержимого таблицы
    static string ExtractTableContent(Table table)
    {
        string tableContent = "";

        foreach (var row in table.Elements<TableRow>())
        {
            tableContent += "<tr>";
            foreach (var cell in row.Elements<TableCell>())
            {
                tableContent += "<td>" + ExtractContent(cell) + "</td>";
            }

            tableContent += "</tr>";
        }

        return tableContent;
    }

    // Получение выравнивания абзаца
    static string GetParagraphAlignment(Paragraph paragraph)
    {
        var alignment = paragraph.ParagraphProperties?.Justification?.Val;
        if (alignment != null)
        {
            switch (alignment)
            {
                case var value when value == JustificationValues.Center:
                    return "center";
                case var value when value == JustificationValues.Right:
                    return "right";
                case var value when value == JustificationValues.Left:
                    return "left";
                default:
                    return "";
            }
        }

        return "";
    }

    // Генерация HTML-кода на основе извлеченного содержимого
    static string GenerateHtmlFromContent(string content)
    {
        return $@"
                <!DOCTYPE html>
                <html lang=""en"">
                <head>
                    <style>
                        table {{
                            font-family: ""Times New Roman"", Times, serif;
                            font-size: 16px;
                        }}
                        p {{
                            font-family: ""Times New Roman"", Times, serif;
                            font-size: 16px;
                            margin: 2px;
                        }}
                        .under_text {{
                            font-size: 10px;
                            text-align: center;
                        }}
                        .under_line {{
                            border-bottom: 1px solid black;
                        }}
                        .rect_bordered {{
                            border: 1px solid black;
                            border-spacing: 0px 0px;
                            width: 25px;
                            height: 25px;
                            font-size: 20px;
                        }}
                        .table_bordered {{
                            border: 1px solid black;
                            border-spacing: 0px 0px;
                            border-collapse: collapse;
                        }}
                        .table_bordered td,
                        th {{
                            border: 1px solid black;
                            border-spacing: 0px 0px;
                            padding: 7px;
                        }}
                        @page {{
                            margin: 0.5cm 1cm 0.5cm 1cm;
                        }}
                        @page :first {{
                            margin: 1cm;
                        }}
                    </style>
                    <title></title>
                </head>
                <body>
                    {content}
                </body>
                </html>
            ";
    }
}