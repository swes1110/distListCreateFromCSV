using System;
using System.IO;
using Microsoft.Exchange.WebServices.Data;
using CsvHelper;
using System.Collections.Generic;
using distListCreateFromCSV.Models;
using distListCreateFromCSV.Mappers;

namespace distListCreateFromCSV
{
    class Program
    {
        static void Main(string[] args)
        {
            //Call function to read CSV file and store in variable
            IEnumerable<contact> contacts = readCsv();
            //Loop thru variable and dispaly data (this is just for testing)
            foreach (var contact in contacts)
            {
                System.Console.Clear();
                System.Console.WriteLine("First Name: " + contact.firstName);
                System.Console.WriteLine("Last Name: " + contact.lastName);
                System.Console.WriteLine("E-Mail Address: " + contact.emailAddress);
            }

            // connectExchange();
        }

        static IEnumerable<contact> readCsv()
        {
            // Read CSV file
            TextReader reader = new StreamReader("templateFile.csv");
            var csvReader = new CsvReader(reader, System.Globalization.CultureInfo.CurrentCulture);
            IEnumerable<contact> contacts = csvReader.GetRecords<contact>();
            return contacts;
        }

        static void connectExchange()
        {
            //Get username from user
            System.Console.WriteLine("Please enter the username for a user with administrative permissions to the exchange server: \r\n");
            string username = Console.ReadLine();
            //Get password
            System.Console.WriteLine("\r\nPlease enter the password for the above user \r\n");
            string password = Console.ReadLine();

            //Create the EWS binding
            ExchangeService service = new ExchangeService();
            //Set credentials
            service.Credentials = new WebCredentials(username,password);
            //Set the service url
            service.Url = new Uri("https://exchange.theabfm.org/EWS/Exchange.asmx");
        }
    }
}