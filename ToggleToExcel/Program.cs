using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Threading.Tasks;
using Toggl;

namespace ToggleToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            String apiKey = TogglToExcel.Properties.Resources.API_key;
            var projectService = new Toggl.Services.ProjectService(apiKey);
            var timeService = new Toggl.Services.TimeEntryService(apiKey);
            var projects = projectService.List();
            
            var timesheets = timeService.List();
            var projectHash = new Dictionary<String,List<TimeEntry>>();
            foreach (var timesheet in timesheets)
            {
                Project project = null;
                foreach (var p in projects)
                {
                    if (p.Id == timesheet.ProjectId)
                    {
                        project = p;
                        break;
                    }
                }
                bool found = false;
                foreach (var p in projectHash.Keys)
                {
                    if (p == project.Name)
                    {
                        found = true;
                        break;
                    }
                }
                if (!found)
                {
                    projectHash.Add(project.Name, new List<TimeEntry>());
                }
                projectHash[project.Name].Add(timesheet);
            }

            foreach (var project in projectHash.Keys)
            {
                Console.WriteLine(project);
                DateTime duration = new DateTime();
                foreach (var timesheet in projectHash[project])
                {
                    DateTime start = Convert.ToDateTime(timesheet.Start);
                    DateTime stop = Convert.ToDateTime(timesheet.Stop);
                    Console.WriteLine("\t" + start.ToShortDateString() + " -> " + start.ToShortTimeString() + " tot " + stop.ToShortTimeString());
                    if(start.DayOfYear == stop.DayOfYear) 
                    {
                        duration += (stop-start);
                    }
                }
                Console.WriteLine("Total: " + duration.ToShortTimeString());
            }
            Console.ReadLine();
        }
    }
}
