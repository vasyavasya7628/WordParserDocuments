using System;
using System.Text;

/// <summary>
/// Summary description for Class1
/// </summary>
public class Class1
{
    public static string Head()
    { 
    StringBuilder builder = new StringBuilder();
    builder.Append(@"<!DOCTYPE html>
        <html lang=""en"">
  <head>
    <meta charset=""UTF-8"">
    
    <title></title>
    <style type=""text/css"">
    table {
    font-family: ""Times New Roman"", Times, serif;
    font-size: 16px;
    }
    p{
    font-family: ""Times New Roman"", Times, serif;
    font-size: 16px;
    margin: 0;
    }
    .under_text {
    
    font-size: 10px;
    text-align: center;
    }
    .under_line {
    border-bottom: 1px solid black;
    }
    .rect_bordered {
    border: 1px solid black;
    border-spacing: 0px 0px;
    width: 25px;
    height: 25px;
    font-size: 20px;
    }
    .table_bordered {
    border: 1px solid black;
    border-spacing: 0px 0px;
    border-collapse: collapse;
    }
    .table_bordered td,
    th {
    border: 1px solid black;
    border-spacing: 0px 0px;
    padding: 3px;
    }
    @page {
    margin: 0.5cm 1cm 0.5cm 1cm;
    }
    @page : first {
    margin: 1cm;
    }
    </style>
  </head>
<body>");
    builder.Replace(@"\\", @"\ \");
        return builder.ToString();
    }


}
