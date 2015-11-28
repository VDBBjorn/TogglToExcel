using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Threading.Tasks;
using Toggl;
using Excel = Microsoft.Office.Interop.Excel;
using TogglToExcel;
using Toggl.QueryObjects; 

namespace ToggleToExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            String apiKey = TogglToExcel.Properties.Resources.APIKey;
            var projectService = new Toggl.Services.ProjectService(apiKey);
            var timeService = new Toggl.Services.TimeEntryService(apiKey);
            var rte = new TimeEntryParams { StartDate = new DateTime(2015,11,9) , EndDate = DateTime.Now};
            var projectHash = GetProjectHash(timeService.List(rte), projectService.List());

            WriteHashToConsole(projectHash);
            WriteToExcel(projectHash);
            
            End();
        }

        private static void WriteToExcel(Dictionary<string, Dictionary<DateTime, TimeSpan>> projectHash)
        {
            //TODO: to dynamic list
            List<String> categorien = new List<String> { "Wiskunde A", "Gedistribueerde toepassingen", "Mobiele toepassingen", "Beveiliging", "Algoritmen", "Systeemanalyse", "Windows", "Masterproef" };
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = false;

            String path = TogglToExcel.Properties.Resources.PlanningFile;
            var workbook = excelApp.Workbooks.Open(path, 0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            var sheet = workbook.Worksheets.get_Item(2);

            Excel.Range excelRange = (Excel.Range)sheet.UsedRange();


            var row = 2;
            foreach (var categorie in categorien)
            {
                if (!projectHash.Keys.Contains(categorie))
                {
                    Console.WriteLine("-" + categorie);
                    row++;
                    continue;
                }
                Console.WriteLine("+" + categorie);
                DateTime date = new DateTime(2015, 11, 9);
                int item = 0;
                for (var col = 3; col <= 84; col++)
                {
                    if (item < projectHash[categorie].Keys.Count && date.Date == projectHash[categorie].Keys.ElementAt(item).Date)
                    {
                        if (excelRange.Cells[row, col].Value == null)
                        {
                            excelRange.Cells[row, col] = projectHash[categorie][date.Date].Duration().TotalHours;
                        }
                        item++;
                    }
                    date = date.AddDays(1);
                }
                row++;
            }
            workbook.Save();
            workbook.Close();
            excelApp.Quit();
        }

        private static Dictionary<String, Dictionary<DateTime, TimeSpan>> GetProjectHash(List<TimeEntry> timesheets, List<Project> projects)
        {
            var projectHash = new Dictionary<String, Dictionary<DateTime,TimeSpan>>();
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
                    projectHash.Add(project.Name, new Dictionary<DateTime,TimeSpan>());
                }
                found = false;
                foreach(var date in projectHash[project.Name].Keys) {
                    if (date == Convert.ToDateTime(timesheet.Start).Date)
                    {
                        found = true;
                        break;
                    }
                }
                var start = Convert.ToDateTime(timesheet.Start);
                var stop = Convert.ToDateTime(timesheet.Stop); 
                if (found)
                {                    
                    projectHash[project.Name][Convert.ToDateTime(timesheet.Start).Date] = projectHash[project.Name][Convert.ToDateTime(timesheet.Start).Date].Add(stop - start);
                }
                else
                {
                    projectHash[project.Name][Convert.ToDateTime(timesheet.Start).Date] = (stop-start);
                }
            }
            return projectHash;
        }

        private static void WriteHashToConsole(Dictionary<String, Dictionary<DateTime,TimeSpan>> projectHash)
        {
            foreach (var project in projectHash.Keys)
            {
                Console.WriteLine(project);
                foreach (var key in projectHash[project].Keys)
                {
                    Console.WriteLine(key.ToShortDateString() + " -> " + projectHash[project][key]);
                }
            }
        }

        private static void End()
        {
            Console.WriteLine("Press any key to exit...");
            Console.ReadLine();
        }
    }
}
