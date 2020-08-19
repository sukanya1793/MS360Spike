﻿using Microsoft.Extensions.Configuration;
using System;
using System.Linq;

namespace GraphTutorial
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine(".NET Core Graph Tutorial\n");

            var appConfig = LoadAppSettings();

            if (appConfig == null)
            {
                Console.WriteLine("Missing or invalid appsettings.json...exiting");
                return;
            }

            var appId = appConfig["appId"];
            var scopesString = appConfig["scopes"];
            var scopes = scopesString.Split(';');

// Initialize the auth provider with values from appsettings.json
            var authProvider = new DeviceCodeAuthProvider(appId, scopes);

// Request a token to sign in the user
            var accessToken = authProvider.GetAccessToken().Result;

            // Initialize Graph client
            GraphHelper.Initialize(authProvider);

// Get signed in user
            var user = GraphHelper.GetMeAsync().Result;
            Console.WriteLine($"Welcome {user.DisplayName}!\n");

            int choice = -1;

            while (choice != 0)
            {
                Console.WriteLine("Please choose one of the following options:");
                Console.WriteLine("0. Exit");
                Console.WriteLine("1. Display access token");
                Console.WriteLine("2. List calendar events");
                Console.WriteLine("3. List license details");

                try
                {
                    choice = int.Parse(Console.ReadLine());
                }
                catch (System.FormatException)
                {
                    // Set to invalid value
                    choice = -1;
                }

                switch (choice)
                {
                    case 0:
                        // Exit the program
                        Console.WriteLine("Goodbye...");
                        break;
                    case 1:
                        // Display access token
                        Console.WriteLine($"Access token: {accessToken}\n");
                        break;
                    case 2:
                        // List the calendar
                        ListCalendarEvents();
                        break;
                    case 3:
                        // List the licenses
                        ListLicenseDetails();
                        break;
                    default:
                        Console.WriteLine("Invalid choice! Please try again.");
                        break;
                }
            }
        }

        static IConfigurationRoot LoadAppSettings()
        {
            var appConfig = new ConfigurationBuilder()
                .AddUserSecrets<Program>()
                .Build();

            // Check for required settings
            if (string.IsNullOrEmpty(appConfig["appId"]) ||
                string.IsNullOrEmpty(appConfig["scopes"]))
            {
                return null;
            }

            return appConfig;
        }

        static string FormatDateTimeTimeZone(Microsoft.Graph.DateTimeTimeZone value)
        {
            // Get the timezone specified in the Graph value
            var timeZone = TimeZoneInfo.FindSystemTimeZoneById(value.TimeZone);
            // Parse the date/time string from Graph into a DateTime
            var dateTime = DateTime.Parse(value.DateTime);

            // Create a DateTimeOffset in the specific timezone indicated by Graph
            var dateTimeWithTZ = new DateTimeOffset(dateTime, timeZone.BaseUtcOffset)
                .ToLocalTime();

            return dateTimeWithTZ.ToString("g");
        }

        static void ListCalendarEvents()
        {
            var events = GraphHelper.GetEventsAsync().Result;

            Console.WriteLine("Events:");

            foreach (var calendarEvent in events)
            {
                Console.WriteLine($"Subject: {calendarEvent.Subject}");
                Console.WriteLine($"  Organizer: {calendarEvent.Organizer.EmailAddress.Name}");
                Console.WriteLine($"  Start: {FormatDateTimeTimeZone(calendarEvent.Start)}");
                Console.WriteLine($"  End: {FormatDateTimeTimeZone(calendarEvent.End)}");
            }
        }


        static void ListLicenseDetails()
        {
            var licenseDetails = GraphHelper.GetLicenseDetailsAsync().Result;

            Console.WriteLine("Service Plans:");

            foreach (var license in licenseDetails)
            {
                foreach (var servicePlan in license.ServicePlans)
                {
                    Console.WriteLine($"  Service Plan ID: {servicePlan.ServicePlanId}");
                    Console.WriteLine($"  Service Plan Name: {servicePlan.ServicePlanName}");
                    Console.WriteLine($"  Provisioning Status: {servicePlan.ProvisioningStatus}");
                    Console.WriteLine("\n------------------------------------------------------");
                }
            }
        }
    }
}