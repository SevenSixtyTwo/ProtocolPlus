using Xceed.Document.NET;
using Xceed.Words.NET;
using System.IO;
using System.Windows.Controls;

public static class WordUtils
{
    public static string CreateNewDocumentFromTemplate(string templatePath, string newFilePath, string newFileName)
    {
        int count = 1;

        if(!Directory.Exists(newFilePath))
        {
            Directory.CreateDirectory(newFilePath);
        }

        string extension = ".docx";
        string newFullPath = Path.Combine(newFilePath, newFileName + extension);

        while (File.Exists(newFullPath))
        {
            string tempFileName = string.Format("{0}({1})", newFileName, count++);
            newFullPath = Path.Combine(newFilePath, tempFileName + extension);
        }

        try
        {
            File.Copy(templatePath, newFullPath, false);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex);
        }

        return newFullPath;
    }

    public static Row GetRowPattern(string filePath, string tableCaption)
    {
        using var document = DocX.Load(filePath);
        var toolsTable = document.Tables.FirstOrDefault(t => t.TableCaption == tableCaption);

        var rowPattern = toolsTable.Rows[1];

        return rowPattern;
    }
    public static Row GetRowPattern(Document protocol, string tableCaption)
    {
        var toolsTable = protocol.Tables.FirstOrDefault(t => t.TableCaption == tableCaption);

        var rowPattern = toolsTable.Rows[1];

        return rowPattern;
    }

    public static void DeleteRowPattern(string filePath, string tableCaption, Row rowPattern)
    {
        using var document = DocX.Load(filePath);
        var toolsTable = document.Tables.FirstOrDefault(t => t.TableCaption == tableCaption);

        toolsTable.RemoveRow(toolsTable.Rows.Count - 1);

        rowPattern.Remove();
        document.Save();
    }
    public static void DeleteRowPattern(Document protocol, string tableCaption, Row rowPattern)
    {
        var toolsTable = protocol.Tables.FirstOrDefault(t => t.TableCaption == tableCaption);

        toolsTable.RemoveRow(toolsTable.Rows.Count - 1);

        rowPattern.Remove();
        protocol.Save();
    }

    public static void CreateRowWithPattern(string filePath, string tableCaption, Row rowPattern, Dictionary<string, string> tagsAndValues)
    {
        using var document = DocX.Load(filePath);
        var toolsTable = document.Tables.FirstOrDefault(t => t.TableCaption == tableCaption);

        AddItemsFromDictToTable(toolsTable, rowPattern, tagsAndValues);

        document.Save();
    }
    public static void CreateRowWithPattern(Document protocol, string tableCaption, Row rowPattern, Dictionary<string, string> tagsAndValues)
    {
        var toolsTable = protocol.Tables.FirstOrDefault(t => t.TableCaption == tableCaption);
        var newItem = toolsTable.InsertRow(rowPattern, toolsTable.RowCount - 1);

        foreach (var item in tagsAndValues)
        {
            newItem.ReplaceText(new StringReplaceTextOptions() { SearchValue = item.Key, NewValue = item.Value });
        }

        protocol.Save();
    }

    public static void AddItemsFromDictToTable(Table table, Row rowPattern, Dictionary<string, string> tagsAndValues)
    {
        var newItem = table.InsertRow(rowPattern, table.RowCount - 1);

        foreach (var item in tagsAndValues)
        {
            newItem.ReplaceText(new StringReplaceTextOptions() { SearchValue = item.Key, NewValue = item.Value });
        }
    }

    public static void CreateRowsFromTemplate(string filePath, string[] tags, string[] values)
    {
        using var document = DocX.Load(filePath);

        var toolsTable = document.Tables.FirstOrDefault(t => t.TableCaption == "TOOLS_TABLE");
        if (toolsTable == null)
        {
            Console.WriteLine("\tError, no TOOLS_TABLE in document");
        }
        else
        {
            if (toolsTable.RowCount > 1)
            {
                var rowPattern = toolsTable.Rows[1];

                AddItemsToTable(toolsTable, rowPattern, tags, values);

                rowPattern.Remove();
            }
        }

        document.Save();
    }
    public static void CreateRowsFromTemplate(Document protocol, string[] tags, string[] values)
    {
        var toolsTable = protocol.Tables.FirstOrDefault(t => t.TableCaption == "TOOLS_TABLE");
        if (toolsTable == null)
        {
            Console.WriteLine("\tError, no TOOLS_TABLE in document");
        }
        else
        {
            if (toolsTable.RowCount > 1)
            {
                var rowPattern = toolsTable.Rows[1];

                AddItemsToTable(toolsTable, rowPattern, tags, values);

                rowPattern.Remove();
            }
        }

        protocol.Save();
    }
    public static void AddItemsToTable(Table table, Row rowPattern, string[] tags, string[] values)
    {
        var newItem = table.InsertRow(rowPattern, table.RowCount - 1);

        for (int i = 0; i < tags.Length; i++)
        {
            newItem.ReplaceText(new StringReplaceTextOptions() { SearchValue = tags[i], NewValue = values[i] });
        }
    }

    public static void ReplaceTagsInFile(string filePath, string tag, string value)
    {
        using (var document = DocX.Load(filePath))
        {
            document.ReplaceText(new StringReplaceTextOptions() { SearchValue = tag, NewValue = value });
            document.Save();
        }
    }
    public static void ReplaceTagsInFile(Document protocol, Dictionary<string, string> tagsAndValues)
    {
        foreach(var item in tagsAndValues) {
            protocol.ReplaceText(new StringReplaceTextOptions() { SearchValue = item.Key, NewValue = item.Value });
        }
        protocol.Save();
    }
}
