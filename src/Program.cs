using System;
using System.Diagnostics;
using System.DirectoryServices;
using System.IO;
using System.Reflection;
using System.Xml;

using DocumentFormat.OpenXml.Spreadsheet;

using SpreadsheetLight;

using Exception = System.Exception;


namespace QueryAddressBook
{
  class Program
  {
    static void Main(string[] args)
    {
      SLDocument excelDoc = null;

      Console.CancelKeyPress += (sender, eventArgs) =>
      {
        if (excelDoc != null)
        {
          var headerStyle = excelDoc.CreateStyle();
          headerStyle.SetFontBold(true);
          headerStyle.SetFontColor(System.Drawing.Color.White);
          headerStyle.SetPatternFill(PatternValues.Solid, System.Drawing.Color.FromArgb(0, 112, 173), System.Drawing.Color.FromArgb(0, 112, 173));

          excelDoc.SetRowStyle(1, headerStyle);
          excelDoc.Filter("A1", "K1");
          excelDoc.FreezePanes(1, 0);
          excelDoc.AutoFitColumn("A1", "K1");

          var outputFile = $"Addressbook_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

          excelDoc.SaveAs(outputFile);
        }

        eventArgs.Cancel = true;
      };


      try
      {
        var searchConditions = ReadSearchConditions("QueryUsers.xml");

        using (var searcher = new DirectorySearcher() { PageSize = 1000 })
        {
          searcher.Filter = "(&(objectClass=user)(objectCategory=person))";
          searcher.PropertiesToLoad.Add("givenName");
          searcher.PropertiesToLoad.Add("sn");
          searcher.PropertiesToLoad.Add("title");
          searcher.PropertiesToLoad.Add("mail");
          searcher.PropertiesToLoad.Add("l");
          searcher.PropertiesToLoad.Add("St");
          searcher.PropertiesToLoad.Add("co");
          searcher.PropertiesToLoad.Add("company");
          searcher.PropertiesToLoad.Add("department");
          searcher.PropertiesToLoad.Add("telephoneNumber");
          searcher.PropertiesToLoad.Add("Mobile");
          searcher.PropertiesToLoad.Add("physicalDeliveryOfficeName");

          using (var results = searcher.FindAll())
          {
            excelDoc = new SLDocument();
            excelDoc.SetCellValue("A1", "Name");
            excelDoc.SetCellValue("B1", "Position");
            excelDoc.SetCellValue("C1", "Business Phone");
            excelDoc.SetCellValue("D1", "Mobile Phone");
            excelDoc.SetCellValue("E1", "E-mail");
            excelDoc.SetCellValue("F1", "Office Location");
            excelDoc.SetCellValue("G1", "Department");
            excelDoc.SetCellValue("H1", "Company Name");
            excelDoc.SetCellValue("I1", "City");
            excelDoc.SetCellValue("J1", "State or Province");
            excelDoc.SetCellValue("K1", "Country or Region");

            int currentRow = 2;

            Console.WriteLine($"Extracting all entries matching the following criteria:");
            Console.WriteLine($"{searchConditions}\n");

            foreach (SearchResult result in results)
            {
              if (searchConditions.IsMatch(result))
              {
                excelDoc.SetCellValue(currentRow, 1, result.GetLastAndFirstNames());
                excelDoc.SetCellValue(currentRow, 2, result.GetProperty("title"));
                excelDoc.SetCellValue(currentRow, 3, result.GetProperty("telephoneNumber"));
                excelDoc.SetCellValue(currentRow, 4, result.GetProperty("Mobile"));
                excelDoc.SetCellValue(currentRow, 5, result.GetProperty("mail"));
                excelDoc.SetCellValue(currentRow, 6, result.GetProperty("physicalDeliveryOfficeName"));
                excelDoc.SetCellValue(currentRow, 7, result.GetProperty("department"));
                excelDoc.SetCellValue(currentRow, 8, result.GetProperty("company"));
                excelDoc.SetCellValue(currentRow, 9, result.GetProperty("l"));
                excelDoc.SetCellValue(currentRow, 10, result.GetProperty("St"));
                excelDoc.SetCellValue(currentRow, 11, result.GetProperty("co"));

                currentRow++;
              }

              Console.Write($"{currentRow - 2} matching entries found\r");
            }
          }
        }

        var headerStyle = excelDoc.CreateStyle();
        headerStyle.SetFontBold(true);
        headerStyle.SetFontColor(System.Drawing.Color.White);
        headerStyle.SetPatternFill(PatternValues.Solid, System.Drawing.Color.FromArgb(0, 112, 173), System.Drawing.Color.FromArgb(0, 112, 173));

        excelDoc.SetRowStyle(1, headerStyle);
        excelDoc.Filter("A1", "K1");
        excelDoc.FreezePanes(1, 0);
        excelDoc.AutoFitColumn("A1", "K1");

        var outputFile = $"Addressbook_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx";

        excelDoc.SaveAs(outputFile);

        Process.Start(new ProcessStartInfo("cmd", $"/c start {Path.GetDirectoryName(Assembly.GetEntryAssembly().Location)}\\{outputFile}"));
      }

      catch (Exception ex)
      {
        Console.Error.WriteLine($"ERROR: {ex.Message}");
      }
    }


    private static void Console_CancelKeyPress(object sender, ConsoleCancelEventArgs e)
    {
      throw new NotImplementedException();
    }


    private static SearchConditions ReadSearchConditions(string file)
    {
      var searchConditions = new SearchConditions();

      var dbFile = new XmlDocument();
      dbFile.Load(file);

      var lastNameNode = dbFile.SelectSingleNode("//SearchConditions/LastName");

      if ((lastNameNode != null) && lastNameNode.NodeType == XmlNodeType.Element)
      {
        searchConditions.LastName = lastNameNode.InnerText;

        if (Enum.TryParse((lastNameNode as XmlElement).GetAttribute("matchMode"), out MatchMode matchMode))
          searchConditions.LastNameMatchMode = matchMode;
      }

      var firstNameNode = dbFile.SelectSingleNode("//SearchConditions/FirstName");

      if ((firstNameNode != null) && firstNameNode.NodeType == XmlNodeType.Element)
      {
        searchConditions.FirstName = firstNameNode.InnerText;

        if (Enum.TryParse((firstNameNode as XmlElement).GetAttribute("matchMode"), out MatchMode matchMode))
          searchConditions.FirstNameMatchMode = matchMode;
      }

      var cityNode = dbFile.SelectSingleNode("//SearchConditions/City");

      if ((cityNode != null) && cityNode.NodeType == XmlNodeType.Element)
      {
        searchConditions.City = cityNode.InnerText;

        if (Enum.TryParse((cityNode as XmlElement).GetAttribute("matchMode"), out MatchMode matchMode))
          searchConditions.CityMatchMode = matchMode;
      }

      var departmentNode = dbFile.SelectSingleNode("//SearchConditions/Department");

      if ((departmentNode != null) && departmentNode.NodeType == XmlNodeType.Element)
      {
        searchConditions.Department = departmentNode.InnerText;

        if (Enum.TryParse((departmentNode as XmlElement).GetAttribute("matchMode"), out MatchMode matchMode))
          searchConditions.DepartmentMatchMode = matchMode;
      }

      var stateOrProvinceNode = dbFile.SelectSingleNode("//SearchConditions/StateOrProvince");

      if ((stateOrProvinceNode != null) && stateOrProvinceNode.NodeType == XmlNodeType.Element)
      {
        searchConditions.StateOrProvince = stateOrProvinceNode.InnerText;

        if (Enum.TryParse((stateOrProvinceNode as XmlElement).GetAttribute("matchMode"), out MatchMode matchMode))
          searchConditions.StateOrProvinceMatchMode = matchMode;
      }

      var countryOrRegionNode = dbFile.SelectSingleNode("//SearchConditions/CountryOrRegion");

      if ((countryOrRegionNode != null) && countryOrRegionNode.NodeType == XmlNodeType.Element)
      {
        searchConditions.CountryOrRegion = countryOrRegionNode.InnerText;

        if (Enum.TryParse((countryOrRegionNode as XmlElement).GetAttribute("matchMode"), out MatchMode matchMode))
          searchConditions.CountryOrRegionMatchMode = matchMode;
      }

      return searchConditions;
    }
  }
}
