using System;
using System.DirectoryServices;
using System.Text;
using System.Text.RegularExpressions;


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

    public MatchMode MobileTelephoneNumberMatchMode { get; set; }

    public string PrimarySmtpAddress { get; set; }

    public MatchMode PrimarySmtpAddressMatchMode { get; set; }

    public string OfficeLocation { get; set; }

    public MatchMode OfficeLocationMatchMode { get; set; }

    public string Department { get; set; }

    public MatchMode DepartmentMatchMode { get; set; }

    public string CompanyName { get; set; }

    public MatchMode CompanyNameMatchMode { get; set; }

    public string City { get; set; }

    public MatchMode CityMatchMode { get; set; }

    public string StateOrProvince { get; set; }

    public MatchMode StateOrProvinceMatchMode { get; set; }

    public string CountryOrRegion { get; set; }

    public MatchMode CountryOrRegionMatchMode { get; set; }


    public bool IsMatch(SearchResult value)
    {
      if (value == null)
        return false;

      bool? lastName = null;
      bool? firstName = null;
      bool? jobTitle = null;
      bool? businessTelephoneNumber = null;
      bool? mobileTelephoneNumber = null;
      bool? primarySmtpAddress = null;
      bool? officeLocation = null;
      bool? department = null;
      bool? companyName = null;
      bool? city = null;
      bool? stateOrProvince = null;
      bool? countryOrRegion = null;

      if (LastName != null)
      {
        var propertyValue = value.GetProperty("sn");

        switch (LastNameMatchMode)
        {
          case MatchMode.IsEqual:

            lastName = propertyValue?.Equals(LastName.Replace(' ', '-'), StringComparison.InvariantCultureIgnoreCase) ?? false;
            break;

          case MatchMode.StartsWith:

            lastName = propertyValue?.StartsWith(LastName.Replace(' ', '-'), StringComparison.InvariantCultureIgnoreCase) ?? false;
            break;

          case MatchMode.Contains:
            lastName = propertyValue?.Contains(LastName.Replace(' ', '-'), StringComparison.InvariantCultureIgnoreCase) ?? false;
            break;

          case MatchMode.RegularExpression:

            lastName = propertyValue != null && Regex.IsMatch(propertyValue, LastName.Replace(' ', '-'), RegexOptions.IgnoreCase);
            break;
        }
      }

      if (FirstName != null)
      {
        var propertyValue = value.GetProperty("givenName");

        switch (FirstNameMatchMode)
        {
          case MatchMode.IsEqual:

            firstName = propertyValue?.Equals(FirstName.Replace(' ', '-'), StringComparison.InvariantCultureIgnoreCase) ?? false;
            break;

          case MatchMode.StartsWith:

            firstName = propertyValue?.StartsWith(FirstName.Replace(' ', '-'), StringComparison.InvariantCultureIgnoreCase) ?? false;
            break;

          case MatchMode.Contains:
            firstName = propertyValue?.Contains(FirstName.Replace(' ', '-'), StringComparison.InvariantCultureIgnoreCase) ?? false;
            break;

          case MatchMode.RegularExpression:

            firstName = propertyValue != null && Regex.IsMatch(propertyValue, FirstName.Replace(' ', '-'), RegexOptions.IgnoreCase);
            break;
        }
      }

      if (JobTitle != null)
      {
        var propertyValue = value.GetProperty("title");

        switch (JobTitleMatchMode)
        {
          case MatchMode.IsEqual:

            jobTitle = propertyValue?.Equals(JobTitle, StringComparison.InvariantCultureIgnoreCase) ?? false;
            break;

          case MatchMode.StartsWith:

            jobTitle = propertyValue?.StartsWith(JobTitle, StringComparison.InvariantCultureIgnoreCase) ?? false;
            break;

          case MatchMode.Contains:

            jobTitle = propertyValue?.Contains(JobTitle, StringComparison.InvariantCultureIgnoreCase) ?? false;
            break;

          case MatchMode.RegularExpression:

            jobTitle = propertyValue != null && Regex.IsMatch(propertyValue, JobTitle, RegexOptions.IgnoreCase);
            break;
        }
      }

      if (BusinessTelephoneNumber != null)
      {
        var propertyValue = value.GetProperty("telephoneNumber");

        switch (BusinessTelephoneNumberMatchMode)
        {
          case MatchMode.IsEqual:

            businessTelephoneNumber = propertyValue?.Equals(BusinessTelephoneNumber, StringComparison.InvariantCultureIgnoreCase) ?? false;
            break;

          case MatchMode.StartsWith:

            businessTelephoneNumber = propertyValue?.StartsWith(BusinessTelephoneNumber, StringComparison.InvariantCultureIgnoreCase) ?? false;
            break;

          case MatchMode.Contains:

            businessTelephoneNumber = propertyValue?.Contains(BusinessTelephoneNumber, StringComparison.InvariantCultureIgnoreCase) ?? false;
            break;

          case MatchMode.RegularExpression:

            businessTelephoneNumber = propertyValue != null && Regex.IsMatch(propertyValue, BusinessTelephoneNumber, RegexOptions.IgnoreCase);
            break;
        }
      }
      /*
      if (!value.Properties.Contains("sn"))
      {
        lastName = false;
      }
      else
      {
        string propertyValue = (string)value.Properties["sn"][0];

        switch (LastNameMatchMode)
        {
          case MatchMode.IsEqual:
            lastName = propertyValue.Equals(LastName, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.StartsWith:
            lastName = propertyValue.StartsWith(LastName, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.Contains:
            lastName = propertyValue.Contains(LastName, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.RegularExpression:
            lastName = Regex.IsMatch(propertyValue, LastName, RegexOptions.IgnoreCase);
            break;
        }
      }
    }

    if (FirstName != null)
    {
      if (!value.Properties.Contains("givenName"))
      {
        firstName = false;
      }
      else
      {
        string propertyValue = (string)value.Properties["givenName"][0];

        switch (FirstNameMatchMode)
        {
          case MatchMode.IsEqual:
            firstName = propertyValue.Equals(FirstName, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.StartsWith:
            firstName = propertyValue.StartsWith(FirstName, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.Contains:
            firstName = propertyValue.Contains(FirstName, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.RegularExpression:
            firstName = Regex.IsMatch(propertyValue, FirstName, RegexOptions.IgnoreCase);
            break;
        }
      }
    }

    if (JobTitle != null)
    {
      if (!value.Properties.Contains("title"))
      {
        jobTitle = false;
      }
      else
      {
        string propertyValue = (string)value.Properties["title"][0];

        switch (JobTitleMatchMode)
        {
          case MatchMode.IsEqual:
            jobTitle = propertyValue.Equals(JobTitle, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.StartsWith:
            jobTitle = propertyValue.StartsWith(JobTitle, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.Contains:
            jobTitle = propertyValue.Contains(JobTitle, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.RegularExpression:
            jobTitle = Regex.IsMatch(propertyValue, JobTitle, RegexOptions.IgnoreCase);
            break;
        }
      }
    }

    if (BusinessTelephoneNumber != null)
    {
      if (!value.Properties.Contains("telephoneNumber"))
      {
        businessTelephoneNumber = false;
      }
      else
      {
        string propertyValue = (string)value.Properties["telephoneNumber"][0];

        switch (BusinessTelephoneNumberMatchMode)
        {
          case MatchMode.IsEqual:
            businessTelephoneNumber = propertyValue.Equals(BusinessTelephoneNumber, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.StartsWith:
            businessTelephoneNumber = propertyValue.StartsWith(BusinessTelephoneNumber, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.Contains:
            businessTelephoneNumber = propertyValue.Contains(BusinessTelephoneNumber, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.RegularExpression:
            businessTelephoneNumber = Regex.IsMatch(propertyValue, BusinessTelephoneNumber, RegexOptions.IgnoreCase);
            break;
        }
      }
    }

    if (MobileTelephoneNumber != null)
    {
      if (!value.Properties.Contains("Mobile"))
      {
        mobileTelephoneNumber = false;
      }
      else
      {
        string propertyValue = (string)value.Properties["Mobile"][0];

        switch (MobileTelephoneNumberMatchMode)
        {
          case MatchMode.IsEqual:
            mobileTelephoneNumber = propertyValue.Equals(MobileTelephoneNumber, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.StartsWith:
            mobileTelephoneNumber = propertyValue.StartsWith(MobileTelephoneNumber, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.Contains:
            mobileTelephoneNumber = propertyValue.Contains(MobileTelephoneNumber, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.RegularExpression:
            mobileTelephoneNumber = Regex.IsMatch(propertyValue, MobileTelephoneNumber, RegexOptions.IgnoreCase);
            break;
        }
      }
    }

    if (PrimarySmtpAddress != null)
    {
      if (!value.Properties.Contains("mail"))
      {
        primarySmtpAddress = false;
      }
      else
      {
        string propertyValue = (string)value.Properties["mail"][0];

        switch (PrimarySmtpAddressMatchMode)
        {
          case MatchMode.IsEqual:
            primarySmtpAddress = propertyValue.Equals(PrimarySmtpAddress, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.StartsWith:
            primarySmtpAddress = propertyValue.StartsWith(PrimarySmtpAddress, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.Contains:
            primarySmtpAddress = propertyValue.Contains(PrimarySmtpAddress, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.RegularExpression:
            primarySmtpAddress = Regex.IsMatch(propertyValue, PrimarySmtpAddress, RegexOptions.IgnoreCase);
            break;
        }
      }
    }

    if (OfficeLocation != null)
    {
      if (!value.Properties.Contains("physicalDeliveryOfficeName"))
      {
        officeLocation = false;
      }
      else
      {
        string propertyValue = (string)value.Properties["physicalDeliveryOfficeName"][0];

        switch (OfficeLocationMatchMode)
        {
          case MatchMode.IsEqual:
            officeLocation = propertyValue.Equals(OfficeLocation, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.StartsWith:
            officeLocation = propertyValue.StartsWith(OfficeLocation, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.Contains:
            officeLocation = propertyValue.Contains(OfficeLocation, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.RegularExpression:
            officeLocation = Regex.IsMatch(propertyValue, OfficeLocation, RegexOptions.IgnoreCase);
            break;
        }
      }
    }

    if (Department != null)
    {
      if (!value.Properties.Contains("department"))
      {
        department = false;
      }
      else
      {
        string propertyValue = (string)value.Properties["department"][0];

        switch (DepartmentMatchMode)
        {
          case MatchMode.IsEqual:
            department = propertyValue.Equals(Department, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.StartsWith:
            department = propertyValue.StartsWith(Department, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.Contains:
            department = propertyValue.Contains(Department, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.RegularExpression:
            department = Regex.IsMatch(propertyValue, Department, RegexOptions.IgnoreCase);
            break;
        }
      }
    }

    if (CompanyName != null)
    {
      if (!value.Properties.Contains("company"))
      {
        companyName = false;
      }
      else
      {
        string propertyValue = (string)value.Properties["company"][0];

        switch (CompanyNameMatchMode)
        {
          case MatchMode.IsEqual:
            companyName = propertyValue.Equals(CompanyName, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.StartsWith:
            companyName = propertyValue.StartsWith(CompanyName, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.Contains:
            companyName = propertyValue.Contains(CompanyName, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.RegularExpression:
            companyName = Regex.IsMatch(propertyValue, CompanyName, RegexOptions.IgnoreCase);
            break;
        }
      }
    }

    if (City != null)
    {
      if (!value.Properties.Contains("l"))
      {
        city = false;
      }
      else
      {
        string propertyValue = (string)value.Properties["l"][0];

        switch (CityMatchMode)
        {
          case MatchMode.IsEqual:
            city = propertyValue.Equals(City, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.StartsWith:
            city = propertyValue.StartsWith(City, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.Contains:
            city = propertyValue.Contains(City, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.RegularExpression:
            city = Regex.IsMatch(propertyValue, City, RegexOptions.IgnoreCase);
            break;
        }
      }
    }

    if (StateOrProvince != null)
    {
      if (!value.Properties.Contains("St"))
      {
        stateOrProvince = false;
      }
      else
      {
        string propertyValue = (string)value.Properties["st"][0];

        switch (StateOrProvinceMatchMode)
        {
          case MatchMode.IsEqual:
            stateOrProvince = propertyValue.Equals(StateOrProvince, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.StartsWith:
            stateOrProvince = propertyValue.StartsWith(StateOrProvince, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.Contains:
            stateOrProvince = propertyValue.Contains(StateOrProvince, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.RegularExpression:
            stateOrProvince = Regex.IsMatch(propertyValue, StateOrProvince, RegexOptions.IgnoreCase);
            break;
        }
      }
    }

    if (CountryOrRegion != null)
    {
      if (!value.Properties.Contains("co"))
      {
        countryOrRegion = false;
      }
      else
      {
        string propertyValue = (string)value.Properties["co"][0];

        switch (CountryOrRegionMatchMode)
        {
          case MatchMode.IsEqual:
            countryOrRegion = propertyValue.Equals(CountryOrRegion, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.StartsWith:
            countryOrRegion = propertyValue.StartsWith(CountryOrRegion, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.Contains:
            countryOrRegion = propertyValue.Contains(CountryOrRegion, StringComparison.InvariantCultureIgnoreCase);
            break;

          case MatchMode.RegularExpression:
            countryOrRegion = Regex.IsMatch(propertyValue, CountryOrRegion, RegexOptions.IgnoreCase);
            break;
        }
      }
    }
    */

      return (lastName != false &&
              firstName != false &&
              jobTitle != false &&
              businessTelephoneNumber != false &&
              mobileTelephoneNumber != false &&
              primarySmtpAddress != false &&
              officeLocation != false &&
              department != false &&
              companyName != false &&
              city != false &&
              stateOrProvince != false &&
              countryOrRegion != false);
    }


    public override string ToString()
    {
      var outStr = new StringBuilder();

      if (LastName != null)
        outStr.AppendLine($"Last Name must {LastNameMatchMode.GetDescription()} '{LastName}'");

      if (FirstName != null)
        outStr.AppendLine($"First Name must {FirstNameMatchMode.GetDescription()} '{FirstName}'");

      if (JobTitle != null)
        outStr.AppendLine($"Job Title must {JobTitleMatchMode.GetDescription()} '{JobTitle}'");

      if (BusinessTelephoneNumber != null)
        outStr.AppendLine($"Business Telephone Number must {BusinessTelephoneNumberMatchMode.GetDescription()} '{BusinessTelephoneNumber}'");

      if (MobileTelephoneNumber != null)
        outStr.AppendLine($"Mobile Telephone Number must {MobileTelephoneNumberMatchMode.GetDescription()} '{MobileTelephoneNumber}'");

      if (PrimarySmtpAddress != null)
        outStr.AppendLine($"Primary SMTP Adress must {PrimarySmtpAddressMatchMode.GetDescription()} '{PrimarySmtpAddress}'");

      if (OfficeLocation != null)
        outStr.AppendLine($"Office Location must {OfficeLocationMatchMode.GetDescription()} '{OfficeLocation}'");

      if (Department != null)
        outStr.AppendLine($"Department must {DepartmentMatchMode.GetDescription()} '{Department}'");

      if (CompanyName != null)
        outStr.AppendLine($"Company Name must {CompanyNameMatchMode.GetDescription()} '{CompanyName}'");

      if (City != null)
        outStr.AppendLine($"City must {CityMatchMode.GetDescription()} '{City}'");

      if (StateOrProvince != null)
        outStr.AppendLine($"State or Province must {StateOrProvinceMatchMode.GetDescription()} '{StateOrProvince}'");

      if (CountryOrRegion != null)
        outStr.AppendLine($"Country or Region must {CountryOrRegionMatchMode.GetDescription()} '{CountryOrRegion}'");

      return outStr.ToString();
    }
  }
}
