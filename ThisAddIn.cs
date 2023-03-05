using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using Serilog;
using System.Diagnostics;

namespace OpenAI
{
    public partial class ThisAddIn
    {
        internal string OpenAIApiKey { get; set; }
        internal string LogFileLocation { get; set; }
        internal ChatGPT3CompletionService ChatGPT3Service { get; set; }

        
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            LogFileLocation = System.Windows.Forms.Application.UserAppDataPath + "\\Netizine-Outlook.log";
            // instantiate and configure logging. Using serilog here, to log to console and a text-file.
            var loggerFactory = new Microsoft.Extensions.Logging.LoggerFactory();
            var loggerConfig = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .WriteTo.Console()
                .WriteTo.File(LogFileLocation, rollingInterval: RollingInterval.Day)
                .CreateLogger();
            loggerFactory.AddSerilog(loggerConfig);
            Serilog.Debugging.SelfLog.Enable(Console.Error);

            OpenAIConfiguration.ApiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY");
            OpenAIApiKey = OpenAIConfiguration.ApiKey;
            if (!string.IsNullOrEmpty(OpenAIApiKey))
            {
                ChatGPT3Service = new ChatGPT3CompletionService();
                Outlook.Accounts accounts = Application.Session.Accounts;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        internal void LogMessage(string data, string callstack, EventLogEntryType eventType = EventLogEntryType.Error)
        {
            try
            {
                Debug.WriteLine(data);
                if (!string.IsNullOrEmpty(callstack))
                {
                    Debug.WriteLine(callstack);
                }
                if (eventType == EventLogEntryType.Error)
                {
                    Log.Error(data);
                }
                else if (eventType == EventLogEntryType.Information)
                {
                    Log.Information(data);
                }
                else if (eventType == EventLogEntryType.Warning)
                {
                    Log.Warning(data);
                }
                else
                {
                    Log.Debug(data);
                }
            }
            catch (Exception)
            {
                // Prevent logging from crashing the plugin
            }
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
