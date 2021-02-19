using System;
using System.IO;
using Microsoft.Exchange.WebServices.Data;
using System.Collections.Generic;
using distListCreateFromCSV.Models;

namespace distListCreateFromCSV
{
    class Program
    {
        static void Main(string[] args)
        {
            //Get username from user
            System.Console.WriteLine("Please enter the username for a user with administrative permissions to the exchange server: \r\n");
            string username = Console.ReadLine();
            //Get password
            System.Console.WriteLine("\r\nPlease enter the password for the above user");
            string password = Console.ReadLine();

            //Create the EWS binding
            ExchangeService service = new ExchangeService();
            //Set credentials
            service.Credentials = new WebCredentials(username,password);
            //Set the service url
            service.Url = new Uri("https://exchange.theabfm.org/EWS/Exchange.asmx");

            //Read CSV file into variable
            TextReader reader = new StreamReader("templateFile.csv");
            CsvHelper.CsvParser parser = new CsvHelper.CsvParser(reader, );
            var csvReader = new CsvHelper.CsvReader(reader);
        }
    }
}