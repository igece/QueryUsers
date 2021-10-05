using System;
using System.ComponentModel;
using System.DirectoryServices;
using System.Linq;
using System.Reflection;


namespace QueryAddressBook
{
  static class Extensions
  {
    public static string GetDescription(this Enum enumValue)
    {
      return enumValue.GetType()
                      .GetMember(enumValue.ToString())
                      .First()
                      .GetCustomAttribute<DescriptionAttribute>()
                      .Description;
    }


    public static string GetProperty(this SearchResult value, string propertyName)
    {
      if (value.Properties.Contains(propertyName))
        return (string)value.Properties[propertyName][0];

      return null;
    }


    public static string GetLastAndFirstNames(this SearchResult value)
    {
      return $"{value.GetProperty("sn")}, {value.GetProperty("givenName")}".Replace('-', ' ');
    }
  }
}
