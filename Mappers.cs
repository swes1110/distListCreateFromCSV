using CsvHelper.Configuration;
using distListCreateFromCSV.Models;
using System.Globalization;

namespace distListCreateFromCSV.Mappers
{
    public sealed class contatMap : ClassMap<contact>
    {
        public void contactMap()
        {
            AutoMap(CultureInfo.InvariantCulture);
            Map(m => m.firstName).Name("firstName");
            Map(m => m.lastName).Name("lastName");
            Map(m => m.emailAddress).Name("emailAddress");
        }
    }
}