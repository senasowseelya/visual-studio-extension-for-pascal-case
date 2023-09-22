using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.ComponentModel.Design;
using System.Globalization;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace PascalCaseExtension
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class ConvertToPascal
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        public const int CommandId = 0x0100;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("7b042abb-78ab-4f62-8adc-0de3f091a326");

        /// <summary>
        /// VS Package that provides this command, not null.
        /// </summary>
        private readonly AsyncPackage package;

        /// <summary>
        /// Initializes a new instance of the <see cref="ConvertToPascal"/> class.
        /// Adds our command handlers for menu (commands must exist in the command table file)
        /// </summary>
        /// <param name="package">Owner package, not null.</param>
        /// <param name="commandService">Command service to add command to, not null.</param>
        private ConvertToPascal(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.MenuItemCallback, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        /// <summary>
        /// Gets the instance of the command.
        /// </summary>
        public static ConvertToPascal Instance
        {
            get;
            private set;
        }

        /// <summary>
        /// Gets the service provider from the owner package.
        /// </summary>
        private Microsoft.VisualStudio.Shell.IAsyncServiceProvider ServiceProvider
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
            // Switch to the main thread - the call to AddCommand in ConvertToPascal's constructor requires
            // the UI thread.
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            Instance = new ConvertToPascal(package, commandService);
        }

        /// <summary>
        /// This function is the callback used to execute the command when the menu item is clicked.
        /// See the constructor to see how the menu item is associated with this function using
        /// OleMenuCommandService service and MenuCommand class.
        /// </summary>
        /// <param name="sender">Event sender.</param>
        /// <param name="e">Event args.</param>
        private async void MenuItemCallback(object sender, EventArgs e)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
            IVsTextView activeView;
            IVsTextLines textLines;
            string selectedText;
            string convertedText = "";
            var textManager = await ServiceProvider.GetServiceAsync(typeof(VsTextManagerClass)) as IVsTextManager;
            textManager.GetActiveView(1, null, out activeView);
            activeView.GetSelection(out int startLine, out int startColumn, out int endLine, out int endColumn);
            if (startLine > endLine)
            {
                (startLine, endLine) = (endLine, startLine);
            }
            if (startColumn > endColumn)
            {
                (startColumn, endColumn) = (endColumn, startColumn);
            }
            activeView.GetBuffer(out textLines);
            activeView.GetSelectedText(out selectedText);
            if (textLines != null && textLines.GetLineText(startLine, startColumn, endLine, endColumn, out selectedText) == 0)
            {
                convertedText = GetPascalCaseText(selectedText);
            }
            IntPtr replacementTextPtr = System.Runtime.InteropServices.Marshal.StringToBSTR(convertedText);
            textLines.ReplaceLines(startLine, startColumn, endLine, endColumn, replacementTextPtr, convertedText.Length, null);
        }

        private string GetPascalCaseText(string selectedText)
        {
            if (string.IsNullOrWhiteSpace(selectedText))
            {
                return string.Empty;
            }
            var pascalCaseWord = new StringBuilder();
            string[] words = selectedText.Split(new[] { " ", "-", "_", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (string word in words)
            {
                pascalCaseWord.Append(char.ToUpper(word[0]));
                pascalCaseWord.Append(word.Substring(1));
            }
            return pascalCaseWord.ToString();
        }
    }
}
