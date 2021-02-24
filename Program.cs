using System;
using System.Net;
using System.Security;
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
            
            //Connect to exchange server
            ExchangeService service = connectExchange();
        }

        static IEnumerable<contact> readCsv()
        {
            // Read CSV file
            TextReader reader = new StreamReader("templateFile.csv");
            var csvReader = new CsvReader(reader, System.Globalization.CultureInfo.CurrentCulture);
            IEnumerable<contact> contacts = csvReader.GetRecords<contact>();
            return contacts;
        }

        static ExchangeService connectExchange()
        {
            //Get username from user
            System.Console.WriteLine("Please enter the username for a user with administrative permissions to the exchange server: \r\n");
            string username = Console.ReadLine();
            //Get password
            SecureString password = GetPasswordFromConsole();

            //Create the EWS binding
            ExchangeService service = new ExchangeService();
            //Set credentials
            service.Credentials = new NetworkCredential(username, password);
            //Set the service url
            service.Url = new Uri("https://exchange.theabfm.org/EWS/Exchange.asmx");
            //Return service object
            return service;
        }

        private static SecureString GetPasswordFromConsole()
        {
            SecureString password = new SecureString();
            bool readingPassword = true;
            Console.Write("Enter password: ");
            while (readingPassword)
            {
                ConsoleKeyInfo userInput = Console.ReadKey(true);
                switch (userInput.Key)
                {
                    case (ConsoleKey.Enter):
                        readingPassword = false;
                        break;
                    case (ConsoleKey.Escape):
                        password.Clear();
                        readingPassword = false;
                        break;
                    case (ConsoleKey.Backspace):
                        if (password.Length > 0)
                        {
                            password.RemoveAt(password.Length - 1);
                            Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                            Console.Write(" ");
                            Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                        }
                        break;
                    default:
                        if (userInput.KeyChar != 0)
                        {
                            password.AppendChar(userInput.KeyChar);
                            Console.Write("*");
                        }
                        break;
                }
            }
            Console.WriteLine();
            password.MakeReadOnly();
            return password;
        }
    }
}