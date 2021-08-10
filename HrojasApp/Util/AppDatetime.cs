using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HrojasApp.Util
{
    public class AppDateTime
    {
        /// <summary>  
        /// Gets or sets the date time format.  
        /// </summary>  
        /// <value>  
        /// The date time format.  
        /// </value>  
        public static string DateTimeFormat { get; set; } = "dd/MM/yyyy H:mm:ss";

        public static string AppTimeZone { get; set; } = "Lima";

        /// <summary>  
        /// Gets or sets the time zone.  
        /// </summary>  
        /// <value>  
        /// The time zone.  
        /// </value>  
        public static TimeZoneInfo TimeZone { get; set; } = TimeZoneInfo.GetSystemTimeZones().First(f => f.DisplayName.Contains(AppTimeZone, StringComparison.CurrentCultureIgnoreCase));

        /// <summary>  
        /// Gets the user local date and time.  
        /// </summary>  
        /// <returns></returns>  
        public static DateTime Now => TimeZoneInfo.ConvertTime(DateTime.UtcNow, TimeZone);


        /// <summary>  
        /// Gets the user local time.  
        /// </summary>  
        /// <returns></returns>  
        public static DateTime Today => TimeZoneInfo.ConvertTime(DateTime.UtcNow, TimeZone).Date;
        /// <summary>  
        /// Gets the UTC time.  
        /// </summary>  
        /// <param name="datetime">The datetime.</param>  
        /// <returns>Get universal date time based on User's Timezone</returns>  
        public static DateTime Convert(DateTime datetime) => TimeZoneInfo.ConvertTime(datetime, TimeZone).ToUniversalTime();
    }
}
