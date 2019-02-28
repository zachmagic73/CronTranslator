using CronExpressionDescriptor;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows;
using Task = System.Threading.Tasks.Task;

namespace CronTranslator
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ShowAllCronExpressionsCommand
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("0f0f857e-3694-466f-a8d8-ffce6bb6abb4");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ShowAllCronExpressionsCommand"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ShowAllCronExpressionsCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ShowAllCronExpressionsCommand Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private IServiceProvider ServiceProvider
        {
            get
            {
                return this.package;
            }
        }

        /// <summary>
        /// Initializes the singleton instance of the command.
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        public static async Task InitializeAsync(AsyncPackage package)
        {
            // Switch to the main thread - the call to AddCommand in ShowAllCronExpressionsCommand's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync((typeof(IMenuCommandService))) as OleMenuCommandService;
            Instance = new ShowAllCronExpressionsCommand(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            DTE dte = (DTE)this.ServiceProvider.GetService(typeof(DTE));
            if (dte.Solution.IsOpen)
            {
                var slnFile = dte.Solution.FullName;
                var directory = Path.GetDirectoryName(slnFile);
                var crons = new List<string>();
                DirSearch(directory, crons);

                if (!crons.Any())
                {
                    VsShellUtilities.ShowMessageBox(
                       this.package,
                       "No timer triggers found in your solution.",
                       "Timer Triggers",
                       OLEMSGICON.OLEMSGICON_INFO,
                       OLEMSGBUTTON.OLEMSGBUTTON_OK,
                       OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
                    return;
                }

                string message = "";
                foreach (string cron in crons)
                {
                    Regex cronRegex = new Regex(@"(\*|((\*\/)?[1-5]?[0-9])) (\*|((\*\/)?[1-5]?[0-9])) (\*|((\*\/)?(1?[0-9]|2[0-3]))) (\*|((\*\/)?([1-9]|[12][0-9]|3[0-1]))) (\*|((\*\/)?([1-9]|1[0-2]))) (\*|((\*\/)?[0-6]))");
                    var cronMatch = cronRegex.Match(cron);
                    string translation = ExpressionDescriptor.GetDescription(cronMatch.Value, new Options()
                    {
                        Verbose = true
                    });
                    message += "\n" + cron.Split('-')[0] + " - " + translation;
                }

                //TODO: Have this command add ReadMe to project
                VsShellUtilities.ShowMessageBox(
                    this.package,
                    message,
                    "Timer Triggers (copied to clipboard)",
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);

                Clipboard.SetText(message);
            }
        }

        static void DirSearch(string sDir, List<string> crons)
        {
            try
            {
                foreach (string d in Directory.GetDirectories(sDir))
                {
                    foreach (string f in Directory.GetFiles(d))
                    {
                        if (f.EndsWith(".cs"))
                        {
                            string content = File.ReadAllText(f);
                            content = content.Replace("\r\n", " ");
                            Regex r = new Regex(@"FunctionName\("".{0,200}(\*|((\*\/)?[1-5]?[0-9])) (\*|((\*\/)?[1-5]?[0-9])) (\*|((\*\/)?(1?[0-9]|2[0-3]))) (\*|((\*\/)?([1-9]|[12][0-9]|3[0-1]))) (\*|((\*\/)?([1-9]|1[0-2]))) (\*|((\*\/)?[0-6]))""");

                            MatchCollection matches = r.Matches(content);
                            int count = matches.Count;

                            if (count > 0)
                            {
                                foreach (Match match in matches)
                                {
                                    Regex funcNameRegex = new Regex(@"FunctionName\("".{0,50}""\)");
                                    var nameMatch = funcNameRegex.Match(match.Value);
                                    Regex cronRegex = new Regex(@"""(\*|((\*\/)?[1-5]?[0-9])) (\*|((\*\/)?[1-5]?[0-9])) (\*|((\*\/)?(1?[0-9]|2[0-3]))) (\*|((\*\/)?([1-9]|[12][0-9]|3[0-1]))) (\*|((\*\/)?([1-9]|1[0-2]))) (\*|((\*\/)?[0-6]))""");
                                    var cronMatch = cronRegex.Match(match.Value);

                                    if (nameMatch.Success && cronMatch.Success)
                                    {
                                        System.Diagnostics.Debug.WriteLine(nameMatch.Value + " " + cronMatch.Value);
                                        crons.Add(nameMatch.Value.Replace("FunctionName(", "").Replace(")", "") + "-" + cronMatch.Value);
                                    }
                                }
                            }
                        }
                    }
                    DirSearch(d, crons);
                }
            }
            catch (System.Exception)
            {
                //ignore
            }
        }
    }
}
