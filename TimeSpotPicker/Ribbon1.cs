using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Windows.Forms;
using System.Globalization;

namespace TimeSpotPicker
{
    [ComVisible(true)]
    public class Ribbon1 : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;
        // Change this to the timezone you want to use for scheduling:
        // "Pacific Standard Time"
        private readonly string TimeZoneForSched = Properties.Settings.Default.AltTimezone;

        private string ordinal (int num)
        {
            string suff = "th";
            int ones = num % 10;
            int tens = (int) Math.Floor(num / 10M) % 10;
            if (tens == 1) {
                suff = "th";
            }
            else
            {
                switch (ones) {
                    case 1 : suff = "st"; break;
                    case 2 : suff = "nd"; break;
                    case 3 : suff = "rd"; break;
                    default: suff = "th"; break;
                }
            }
            return String.Format("{0}{1}", num, suff);
        }

        public Ribbon1()
        {
            
        }

        public void copySpot(Office.IRibbonControl control)
        {
            // Get selected calendar date
            // Thanks to https://stackoverflow.com/questions/25040715/outlook-addin-get-current-selected-calendar-date
            Outlook.Application application = new Outlook.Application();
            Outlook.Explorer explorer = application.ActiveExplorer();
            Outlook.Folder folder = explorer.CurrentFolder as Outlook.Folder;
            Outlook.View view = explorer.CurrentView as Outlook.View;

            if (view.ViewType == Outlook.OlViewType.olCalendarView)
            {
                Outlook.CalendarView calView = view as Outlook.CalendarView;
                DateTime startTime = calView.SelectedStartTime;
                DateTime endTime = calView.SelectedEndTime;
                string timeSpot = "Something went wrong";
                try
                {
                    if (Properties.Settings.Default.UseAltTimezone == true)
                    {
                        TimeZoneInfo altTimezone = TimeZoneInfo.FindSystemTimeZoneById(TimeZoneForSched);
                        startTime = TimeZoneInfo.ConvertTime(startTime, altTimezone);
                        endTime = TimeZoneInfo.ConvertTime(endTime, altTimezone);
                    }

                    if (startTime.Day == endTime.Day)
                    {
                        timeSpot = startTime.DayOfWeek.ToString().Substring(0, 3) + " ";
                        timeSpot += startTime.ToString("MMM", CultureInfo.InvariantCulture) + " ";
                        timeSpot += ordinal(startTime.Day) + ", ";
                        timeSpot += startTime.ToString(@"hh\:mm");
                        timeSpot += startTime.ToString("tt", CultureInfo.InvariantCulture).ToLower() + " - ";
                        timeSpot += endTime.ToString(@"hh\:mm");
                        timeSpot += endTime.ToString("tt", CultureInfo.InvariantCulture).ToLower();
                    }
                    else
                        timeSpot = "Different days? Sure?";
                }
                catch (TimeZoneNotFoundException)
                {
                    timeSpot = "Unable to find the " + TimeZoneForSched + " zone in the registry.";
                }
                catch (InvalidTimeZoneException)
                {
                    timeSpot = "Looks like " + TimeZoneForSched + " is not a valid timezone.";
                }
                Debug.Print(timeSpot);
                //Clipboard.SetText(timeSpot);
                Clipboard.SetDataObject(timeSpot, false, 0, 0);
                Marshal.ReleaseComObject(application);
                Marshal.ReleaseComObject(view);
                application = null;
                view = null;
                return;
            }
        }


        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("TimeSpotPicker.Ribbon1.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
