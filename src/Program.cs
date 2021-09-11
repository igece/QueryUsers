using System;
using System.Xml;

using DocumentFormat.OpenXml.Spreadsheet;

using Microsoft.Office.Interop.Outlook;

using SpreadsheetLight;


namespace QueryAddressBook
{
  class Program
  {
    static void Main(string[] args)
    {
      var searchConditions = ReadSearchConditions("QueryAddressBook.xml");

      var app = new Application();
      var addressList = app.Session.GetGlobalAddressList();

      if (addressList == null)
      {
        Console.Error.WriteLine("ERROR: Unable to obtain address list. Is Outlook installed?");
        return;
      }

      using var excelDoc = new SLDocument();
      excelDoc.SetCellValue("A1", "Last Name");
      excelDoc.SetCellValue("B1", "First Name");
      excelDoc.SetCellValue("C1", "Position");
      excelDoc.SetCellValue("D1", "Business Phone");
      excelDoc.SetCellValue("E1", "Mobile Phone");
      excelDoc.SetCellValue("F1", "E-mail");
      excelDoc.SetCellValue("G1", "Office Location");
      excelDoc.SetCellValue("H1", "Department");
      excelDoc.SetCellValue("I1", "Company Name");
      excelDoc.SetCellValue("J1", "City");
      excelDoc.SetCellValue("K1", "State or Province");

      int currentRow = 2;
      int currentEntry = 0;
      int totalEntries = addressList.AddressEntries.Count;

      foreach (AddressEntry addressEntry in addressList.AddressEntries)
      {
        if (addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeUserAddressEntry ||
            addressEntry.AddressEntryUserType == OlAddressEntryUserType.olExchangeRemoteUserAddressEntry)
        {
          var exchangeUser = addressEntry.GetExchangeUser();

          if (searchConditions.IsMatch(exchangeUser))
          {
            excelDoc.SetCellValue(currentRow, 1, exchangeUser.LastName);
            excelDoc.SetCellValue(currentRow, 2, exchangeUser.FirstName);
            excelDoc.SetCellValue(currentRow, 3, exchangeUser.JobTitle);
            excelDoc.SetCellValue(currentRow, 4, exchangeUser.BusinessTelephoneNumber);
            excelDoc.SetCellValue(currentRow, 5, exchangeUser.MobileTelephoneNumber);
            excelDoc.SetCellValue(currentRow, 6, exchangeUser.PrimarySmtpAddress);
            excelDoc.SetCellValue(currentRow, 7, exchangeUser.OfficeLocation);
            excelDoc.SetCellValue(currentRow, 8, exchangeUser.Department);
            excelDoc.SetCellValue(currentRow, 9, exchangeUser.CompanyName);
            excelDoc.SetCellValue(currentRow, 10, exchangeUser.City);
            excelDoc.SetCellValue(currentRow, 11, exchangeUser.StateOrProvince);
            currentRow++;
          }

          Console.Write($"{(double)currentEntry++ / totalEntries * 100:F02}% completed\r");
        }
      }

      var headerStyle = excelDoc.CreateStyle();
      headerStyle.SetFontBold(true);
      headerStyle.SetFontColor(System.Drawing.Color.White);
      headerStyle.SetPatternFill(PatternValues.Solid, System.Drawing.Color.FromArgb(0, 112, 173), System.Drawing.Color.FromArgb(0, 112, 173));

      excelDoc.AutoFitColumn("A1", "K1");
      excelDoc.SetRowStyle(1, headerStyle);
      excelDoc.Filter("A1", "K1");
      excelDoc.FreezePanes(1, 0);

      excelDoc.SaveAs($"Addressbook_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx");
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

      return searchConditions; 
    }
  }
}
