using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32.TaskScheduler;
using Newtonsoft.Json;

/*
    Author : Aman R. Waghmare
    Date   : 16-04-2023
    What it does : reads json file deserialize it and schedule a Myemailsender.exe file to trigger at given date and time.
 */
namespace Myemailsender
{
    [Serializable]
    internal class DbBackupdata
    {
        public List<string> Repeat { get; set; }
        public string TimeOfBackup { get; set; }
        public string OverwriteAfter { get; set; }
        public string EveryHour { get; set; }
    }

    internal class Program
    {
        static void Main(string[] args)
        {
            Program.scheduletask();
        }


        public static void scheduletask()
        {
            Console.WriteLine("starting task scheduler...");

            Console.WriteLine("reading json...");
            // read json file
            string filename = "E:\\DbBackupdemo.json";

            if (File.Exists(filename))
            {
                Console.WriteLine("file exist..");

                string jsonf = File.ReadAllText(filename);

                // deserialization
                DbBackupdata myobj = JsonConvert.DeserializeObject<DbBackupdata>(jsonf);
                
                // user input from json file
                int reapeat_hours = Convert.ToInt32(myobj.EveryHour);

                Console.WriteLine("scheduling emailsender.exe for given Date and timing..");

                // Extract the hour and minute values from the TimeOfBackup property
                var timeParts = myobj.TimeOfBackup.Split(':');
                int hour = int.Parse(timeParts[0]);
                int minute = int.Parse(timeParts[1]);

                // path of emailsender.exe
                string event_execute = @"C:\Users\anujb\source\repos\Myemailsender\Myemailsender\bin\Debug\net6.0\Myemailsender.exe";
              
                // Create a new task scheduler
                TaskService ts = new TaskService();

                // Retrieve the existing "\Rapidigital_events" folder, or create it if it doesn't exist
                TaskFolder Folder = ts.RootFolder.SubFolders.FirstOrDefault(f => f.Name == "Rapidigital_events");
                if (Folder == null)
                {
                    Folder = ts.RootFolder.CreateFolder("Rapidigital_events");
                }

                // Create a new task
                TaskDefinition td = ts.NewTask();

                // Set the task properties
                td.RegistrationInfo.Description = "custome_task of emails";
                td.RegistrationInfo.Author = "Rapidigital_services";



                // Get the value of the "Repeat" property and split it into parts
                var repeatParts = myobj.Repeat[0].Split(',');


                // Set the date based on the value of "Repeat"
                DateTime date = DateTime.Today;

                if (repeatParts[0].ToLower() == "daily")
                {
                    // Create a new daily trigger
                    DailyTrigger dailyTrigger = new DailyTrigger();


                    // Set the start boundary to the same time each day as per user input
                    dailyTrigger.StartBoundary = new DateTime(date.Year, date.Month, date.Day, hour, minute, 0);

                    // Add the trigger to the task definition
                    td.Triggers.Add(dailyTrigger);

                    // days interval every 1 day
                    dailyTrigger.DaysInterval = 1;


                    // Set the repetition interval to repeat every 3 hours
                    dailyTrigger.Repetition.Interval = TimeSpan.FromHours(reapeat_hours);
                    // repeat only 1 trigger/24 hours
                    dailyTrigger.Repetition.Duration = TimeSpan.FromDays(1);
                }

                else if (repeatParts[0].ToLower() == "weekly")
                {
                    // set trigger for weekly backup
                    WeeklyTrigger weeklyTrigger = new WeeklyTrigger();
                    //
                    // Set the start boundary to the same time each day as per user input
                    weeklyTrigger.StartBoundary = new DateTime(date.Year, date.Month, date.Day, hour, minute, 0);

                    // Set the repetition interval to repeat every 3 hours
                    weeklyTrigger.Repetition.Interval = TimeSpan.FromHours(reapeat_hours);

                    // set days of week for trigger based on repeat value
                    for (int i = 1; i < repeatParts.Length; i++)
                    {
                        switch (repeatParts[i].Trim().ToLower())
                        {
                            case "monday":
                                weeklyTrigger.DaysOfWeek |= DaysOfTheWeek.Monday;
                                break;

                            case "tuesday":
                                weeklyTrigger.DaysOfWeek |= DaysOfTheWeek.Tuesday;
                                break;

                            case "wednesday":
                                weeklyTrigger.DaysOfWeek |= DaysOfTheWeek.Wednesday;
                                break;

                            case "thursday":
                                weeklyTrigger.DaysOfWeek |= DaysOfTheWeek.Thursday;
                                break;

                            case "friday":
                                weeklyTrigger.DaysOfWeek |= DaysOfTheWeek.Friday;
                                break;

                            case "saturday":
                                weeklyTrigger.DaysOfWeek |= DaysOfTheWeek.Saturday;
                                break;

                            case "sunday":
                                weeklyTrigger.DaysOfWeek |= DaysOfTheWeek.Sunday;
                                break;

                            default:
                                Console.WriteLine("missing day");
                                continue;
                        }
                    }
                    // Add the weekly trigger to the task definition
                    td.Triggers.Add(weeklyTrigger);
                }
                else
                {
                    // set trigger for monthly backup
                    MonthlyTrigger monthlyTrigger = new MonthlyTrigger();
                    monthlyTrigger.StartBoundary = new DateTime(date.Year, date.Month, DateTime.Now.Day, hour, minute, 0);

                    // get Repeat value as string
                    string repeatValue = Convert.ToString(myobj.Repeat[0]);

                    // Split the Repeat value by comma and trim each part
                    string[] repeatPartsm = repeatValue.Split(',').Select(p => p.Trim()).ToArray();

                    // get the day(s) of the month specified in the Repeat value
                    int[] daysOfMonth = new int[repeatPartsm.Length-1];


                    for (int i = 1; i < repeatPartsm.Length; i++)
                    {
                        daysOfMonth[i - 1] = Convert.ToInt32(repeatParts[i]);  // not in correct format error
                    }

                    Console.WriteLine(Convert.ToInt32(daysOfMonth[0]));
                    
                    monthlyTrigger.DaysOfMonth = daysOfMonth;
                    monthlyTrigger.MonthsOfYear = MonthsOfTheYear.AllMonths;
                    // Add the monthly trigger to the task definition
                    td.Triggers.Add(monthlyTrigger);
                }


                // Print the date to the console
       //         Console.WriteLine($"Date is: {date.ToString("dd/MM/yyyy")}");


                // action to be trigger  this is to lounch executeble file
                td.Actions.Add(new ExecAction(event_execute, null, null));


                // Register the task_name in the "\Rapidigital_events" folder of the task scheduler
                Folder.RegisterTaskDefinition("Auto_Emails", td);

                Console.WriteLine("email sender Scheduled successfully..");
                Console.ReadKey();
            }
        }
    }
}
