using System.Text.RegularExpressions;

using Microsoft.Office.Interop.Outlook;


namespace QueryAddressBook
{
  class SearchConditions
  {
    public string LastName { get; set; }

    public MatchMode LastNameMatchMode { get; set; }

    public string FirstName { get; set; }

    public MatchMode FirstNameMatchMode { get; set; }

    public string JobTitle { get; set; }

    public MatchMode JobTitleMatchMode { get; set; }

    public string BusinessTelephoneNumber { get; set; }

    public MatchMode BusinessTelephoneNumberMatchMode { get; set; }

    public string MobileTelephoneNumber { get; set; }

    public string PrimarySmtpAddress { get; set; }

    public string OfficeLocation { get; set; }

    public string Department { get; set; }

    public string CompanyName { get; set; }

    public string City { get; set; }

    public MatchMode CityMatchMode { get; set; }

    public string StateOrProvince { get; set; }


    public bool IsMatch(ExchangeUser value)
    {
      if (value == null)
        return false;

      bool? lastName = null;
      bool? firstName = null;
      bool? jobTitle = null;
      bool? businessTelephoneNumber = null;
      bool? city = null;

      if (LastName != null)
      {
        if (value.LastName == null)
        {
          lastName = false;
        }
        else
        {
          switch (LastNameMatchMode)
          {
            case MatchMode.IsEqual:
              lastName = value.LastName.Equals(LastName);
              break;

            case MatchMode.StartsWith:
              lastName = value.LastName.StartsWith(LastName);
              break;

            case MatchMode.Contains:
              lastName = value.LastName.Contains(LastName);
              break;

            case MatchMode.Regexp:
              lastName = Regex.IsMatch(value.LastName, LastName);
              break;
          }
        }
      }

      if (FirstName != null)
      {
        if (value.FirstName == null)
        {
          firstName = false;
        }
        else
        {
          switch (FirstNameMatchMode)
          {
            case MatchMode.IsEqual:
              firstName = value.FirstName.Equals(FirstName);
              break;

            case MatchMode.StartsWith:
              firstName = value.FirstName.StartsWith(FirstName);
              break;

            case MatchMode.Contains:
              firstName = value.FirstName.Contains(FirstName);
              break;

            case MatchMode.Regexp:
              firstName = Regex.IsMatch(value.FirstName, FirstName);
              break;
          }
        }
      }

      if (JobTitle != null)
      {
        if (value.JobTitle == null)
        {
          jobTitle = false;
        }
        else
        {
          switch (JobTitleMatchMode)
          {
            case MatchMode.IsEqual:
              jobTitle = value.JobTitle.Equals(JobTitle);
              break;

            case MatchMode.StartsWith:
              jobTitle = value.JobTitle.StartsWith(JobTitle);
              break;

            case MatchMode.Contains:
              jobTitle = value.JobTitle.Contains(JobTitle);
              break;

            case MatchMode.Regexp:
              jobTitle = Regex.IsMatch(value.JobTitle, JobTitle);
              break;
          }
        }
      }

      if (BusinessTelephoneNumber != null)
      {
        if (value.BusinessTelephoneNumber == null)
        {
          businessTelephoneNumber = false;
        }
        else
        {
          switch (BusinessTelephoneNumberMatchMode)
          {
            case MatchMode.IsEqual:
              businessTelephoneNumber = value.BusinessTelephoneNumber.Equals(BusinessTelephoneNumber);
              break;

            case MatchMode.StartsWith:
              businessTelephoneNumber = value.BusinessTelephoneNumber.StartsWith(BusinessTelephoneNumber);
              break;

            case MatchMode.Contains:
              businessTelephoneNumber = value.BusinessTelephoneNumber.Contains(BusinessTelephoneNumber);
              break;

            case MatchMode.Regexp:
              businessTelephoneNumber = Regex.IsMatch(value.BusinessTelephoneNumber, BusinessTelephoneNumber);
              break;
          }
        }
      }

      if (City != null)
      {
        if (value.City == null)
        {
          city = false;
        }
        else
        {
          switch (CityMatchMode)
          {
            case MatchMode.IsEqual:
              city = value.City.Equals(City);
              break;

            case MatchMode.StartsWith:
              city = value.City.StartsWith(City);
              break;

            case MatchMode.Contains:
              city = value.City.Contains(City);
              break;

            case MatchMode.Regexp:
              city = Regex.IsMatch(value.City, City);
              break;
          }
        }
      }

      return (lastName != false &&
              firstName != false &&
              jobTitle != false &&
              businessTelephoneNumber != false &&
              city != false);
    }
  }
}
