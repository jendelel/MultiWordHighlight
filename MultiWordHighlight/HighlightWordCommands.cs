using System;
using System.ComponentModel.Design;
using System.Globalization;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Text;
using Microsoft.VisualStudio.Text.Operations;

namespace MultiWordHighlight
{
    /// <summary>
    /// Command handler
    /// </summary>
    internal sealed class HighlightWordCommands
    {
        /// <summary>
        /// Command ID.
        /// </summary>
        const int cmdidAddColumnGuide = 0x0100;
        const int cmdidRemoveColumnGuide = 0x0101;
        const int cmdidChooseGuideColor = 0x0102;
        const int cmdidRemoveAllColumnGuides = 0x0103;

        /// <summary>
        /// Command menu group (command set GUID).
        /// </summary>
        public static readonly Guid CommandSet = new Guid("664e7751-1dbe-4bef-8bd6-e5450f999df5");


        /// <summary>  
        /// VS Package that provides this command, not null.  
        /// </summary>  
        private readonly Package package;

        OleMenuCommand _toggleWordCommand;

        public ITextStructureNavigatorSelectorService TextStructureNavigatorSelectorService
        {
            get; set;
        }

        /// <summary>  
        /// Initializes the singleton instance of the command.  
        /// </summary>  
        /// <param name="package">Owner package, not null.</param>  
        public static void Initialize(Package package, ITextStructureNavigatorSelectorService textStructureNavigatorSelector)
        {
            Instance = new HighlightWordCommands(package, textStructureNavigatorSelector);
        }

        /// <summary>  
        /// Gets the instance of the command.  
        /// </summary>  
        public static HighlightWordCommands Instance
        {
            get;
            private set;
        }

        /// <summary>  
        /// Initializes a new instance of the <see cref="ColumnGuideCommands"/> class.  
        /// Adds our command handlers for menu (commands must exist in the command   
        /// table file)  
        /// </summary>  
        /// <param name="package">Owner package, not null.</param>  
        private HighlightWordCommands(Package package, ITextStructureNavigatorSelectorService textStructureNavigatorSelector)
        {
            if (package == null)
            {
                throw new ArgumentNullException("package");
            }

            this.package = package;
            this.TextStructureNavigatorSelectorService = textStructureNavigatorSelector;

            // Add our command handlers for menu (commands must exist in the .vsct file)  

            OleMenuCommandService commandService =
                this.ServiceProvider.GetService(typeof(IMenuCommandService))
                    as OleMenuCommandService;
            if (commandService != null)
            {
                // Add guide  
                _toggleWordCommand =
                    new OleMenuCommand(AddWordExecuted, null,
                                       AddWordBeforeQueryStatus,
                                       new CommandID(HighlightWordCommands.CommandSet,
                                                     cmdidAddColumnGuide));
                _toggleWordCommand.ParametersDescription = "<word>";
                commandService.AddCommand(_toggleWordCommand);

                // Remove all  
                commandService.AddCommand(
                    new MenuCommand(RemoveAllWordsExecuted,
                                    new CommandID(HighlightWordCommands.CommandSet,
                                                  cmdidRemoveAllColumnGuides)));
            }
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

        private void AddWordBeforeQueryStatus(object sender, EventArgs e)
        {
            string currentWord = GetCurrentEditorWord();
            _toggleWordCommand.Enabled =
                HighlightWordsSettingsManager.CanToggleWord(currentWord);
        }

        private string GetCurrentEditorWord()
        {
            IVsTextView view = GetActiveTextView();
            if (view == null)
            {
                return null;
            }

            try
            {
                IWpfTextView textView = GetTextViewFromVsTextView(view);
                if (textView == null || TextStructureNavigatorSelectorService == null) return null;
                ITextStructureNavigator textStructureNavigator = TextStructureNavigatorSelectorService.GetTextStructureNavigator(textView.TextBuffer);
                string column = GetCaretWord(textView, textStructureNavigator);

                // Note: GetCaretColumn returns 0-based positions. Guidelines are 1-based  
                // positions.  
                // However, do not subtract one here since the caret is positioned to the  
                // left of  
                // the given column and the guidelines are positioned to the right. We  
                // want the  
                // guideline to line up with the current caret position. e.g. When the  
                // caret is  
                // at position 1 (zero-based), the status bar says column 2. We want to  
                // add a  
                // guideline for column 1 since that will place the guideline where the  
                // caret is.  
                return column;
            }
            catch (InvalidOperationException)
            {
                return null;
            }
        }

        /// <summary>  
        /// Find the active text view (if any) in the active document.  
        /// </summary>  
        /// <returns>The IVsTextView of the active view, or null if there is no active  
        /// document or the  
        /// active view in the active document is not a text view.</returns>  
        private IVsTextView GetActiveTextView()
        {
            IVsMonitorSelection selection =
                this.ServiceProvider.GetService(typeof(IVsMonitorSelection))
                                                    as IVsMonitorSelection;
            object frameObj = null;
            ErrorHandler.ThrowOnFailure(
                selection.GetCurrentElementValue(
                    (uint)VSConstants.VSSELELEMID.SEID_DocumentFrame, out frameObj));

            IVsWindowFrame frame = frameObj as IVsWindowFrame;
            if (frame == null)
            {
                return null;
            }

            return GetActiveView(frame);
        }

        private static IVsTextView GetActiveView(IVsWindowFrame windowFrame)
        {
            if (windowFrame == null)
            {
                throw new ArgumentException("windowFrame");
            }

            object pvar;
            ErrorHandler.ThrowOnFailure(
                windowFrame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out pvar));

            IVsTextView textView = pvar as IVsTextView;
            if (textView == null)
            {
                IVsCodeWindow codeWin = pvar as IVsCodeWindow;
                if (codeWin != null)
                {
                    ErrorHandler.ThrowOnFailure(codeWin.GetLastActiveView(out textView));
                }
            }
            return textView;
        }

        private static IWpfTextView GetTextViewFromVsTextView(IVsTextView view)
        {

            if (view == null)
            {
                throw new ArgumentNullException("view");
            }

            IVsUserData userData = view as IVsUserData;
            if (userData == null)
            {
                throw new InvalidOperationException();
            }

            object objTextViewHost;
            if (VSConstants.S_OK
                   != userData.GetData(Microsoft.VisualStudio
                                                .Editor
                                                .DefGuidList.guidIWpfTextViewHost,
                                       out objTextViewHost))
            {
                throw new InvalidOperationException();
            }

            IWpfTextViewHost textViewHost = objTextViewHost as IWpfTextViewHost;
            if (textViewHost == null)
            {
                throw new InvalidOperationException();
            }

            return textViewHost.TextView;
        }

        /// <summary>  
        /// Given an IWpfTextView, find the position of the caret and report its column  
        /// number. The column number is 0-based  
        /// </summary>  
        /// <param name="textView">The text view containing the caret</param>  
        /// <returns>The column number of the caret's position. When the caret is at the  
        /// leftmost column, the return value is zero.</returns>  
        private static string GetCaretWord(IWpfTextView textView, ITextStructureNavigator textStructureNavigator)
        {
            var caretPosition = textView.Caret.Position;
            var sourceBuffer = textView.TextBuffer;
            SnapshotPoint? point = caretPosition.Point.GetPoint(sourceBuffer, caretPosition.Affinity);
            if (!point.HasValue)
                return null;

            var snapshotPoint = point.Value;
            TextExtent word = textStructureNavigator.GetExtentOfWord(snapshotPoint);
            var foundWord = true;
            if (!HighlightWordTagger.WordExtentIsValid(snapshotPoint, word))
            {
                //Before we retry, make sure it is worthwhile   
                if (word.Span.Start != snapshotPoint
                     || snapshotPoint == snapshotPoint.GetContainingLine().Start
                     || char.IsWhiteSpace((snapshotPoint - 1).GetChar()))
                {
                    foundWord = false;
                }
                else
                {
                    // Try again, one character previous.    
                    //If the caret is at the end of a word, pick up the word.  
                    word = textStructureNavigator.GetExtentOfWord(snapshotPoint - 1);

                    //If the word still isn't valid, we're done   
                    if (!HighlightWordTagger.WordExtentIsValid(snapshotPoint, word))
                        foundWord = false;
                }
            }

            if (!foundWord)
            {
                return null;
            }
            return word.Span.GetText();
        }

        /// <summary>  
        /// Determine the applicable word for an add or remove command.  
        /// The word is parsed from command arguments, if present. Otherwise  
        /// the current position of the caret is used to determine the word.  
        /// </summary>  
        /// <param name="e">Event args passed to the command handler.</param>  
        /// <returns>The column number. May be negative to indicate the column number is  
        /// unavailable.</returns>  
        /// <exception cref="ArgumentException">The column number parsed from event args  
        /// was not a valid integer.</exception>  
        private string GetApplicableWord(EventArgs e)
        {
            var inValue = ((OleMenuCmdEventArgs)e).InValue as string;
            if (!string.IsNullOrEmpty(inValue) && inValue.IndexOf(' ') < 0)
            {
                return inValue;
            }

            return GetCurrentEditorWord();
        }

        /// <summary>  
        /// This function is the callback used to execute a command when the a menu item  
        /// is clicked. See the Initialize method to see how the menu item is associated  
        /// to this function using the OleMenuCommandService service and the MenuCommand  
        /// class.  
        /// </summary>  
        private void AddWordExecuted(object sender, EventArgs e)
        {
            string word = GetApplicableWord(e);
            if (!string.IsNullOrEmpty(word))
            {
                HighlightWordsSettingsManager.ToggleWord(word);
            }
        }

        private void RemoveAllWordsExecuted(object sender, EventArgs e)
        {
            HighlightWordsSettingsManager.RemoveAllWords();
        }

    }
}