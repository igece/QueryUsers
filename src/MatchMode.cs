using System.ComponentModel;

namespace QueryAddressBook
{
  enum MatchMode
  {
    [Description("be equal to")]
    IsEqual,
    
    [Description("start with")]
    StartsWith,

    [Description("contain")]
    Contains,

    [Description("match regular expression")]
    RegularExpression
  }
}