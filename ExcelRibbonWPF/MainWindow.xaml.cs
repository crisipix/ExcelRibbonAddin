using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelRibbonWPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private Excel.Application _xlApp;
        public MainWindow(Excel.Application xlApp)
        {
            _xlApp = xlApp;
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            if (_xlApp != null)
            {
                Excel.Worksheet workSheet = _xlApp.ActiveWorkbook.Worksheets[1];
                workSheet.Cells[1, "A"] = "ID Number";
                workSheet.Cells[1, "B"] = "Current Balance";

                var accounts = Enumerable.Range(1, 100)
                                         .Select((x, y) => new { Id = 100000 + y, Balance = new Random().Next(0, 300 * x) });
                var index = 2;
                foreach (var acct in accounts)
                {
                    workSheet.Cells[index, "A"] = acct.Id;
                    workSheet.Cells[index, "B"] = acct.Balance;
                    index++;
                }

                // Call to AutoFormat in Visual C# 2010. This statement replaces the 
                // two calls to AutoFit.
                workSheet.Range["A1", "B3"].AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);

                Excel.Worksheet ws = _xlApp.ActiveSheet;
                Excel.Range range = ws.Range[ws.Cells[1, 1], ws.Cells[(2), (5)]]; // row 1 col 1 to row 2 col 5
                ws.Names.Add("ChrisRange", range);//"Sheet1!ChrisRange"

                //https://msdn.microsoft.com/en-us/library/microsoft.office.tools.excel.namedrange.aspx
                //https://msdn.microsoft.com/en-us/library/7zte17ya.aspx
                //https://msdn.microsoft.com/en-us/library/bb386091.aspx
                // find the range "Sheet1!ChrisRange" and format it. 
                foreach (Excel.Name n in ws.Names)
                {
                    Console.WriteLine(n.Name);
                    Excel.Range r = n.RefersToRange;
                    r.AutoFormat(Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);
                }
            }
        }

        private void Button_Click_Highlight(object sender, RoutedEventArgs e)
        {
            Excel.Worksheet ws = _xlApp.ActiveSheet;

            // find the range "Sheet1!ChrisRange" and format it. 
            foreach (Excel.Name n in ws.Names)
            {
                Console.WriteLine(n.Name);
                var chrisRangeName = $"{ws.Name}!ChrisRange";
                if (n.Name == chrisRangeName)
                {
                    Excel.Range range = n.RefersToRange;
                    range.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                    range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                    //range.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    //range.Borders.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(153, 153, 153));

                    //https://msdn.microsoft.com/en-us/library/bb386091.aspx
                    //Search for text in worksheet ranges.
                }

            }

        }
        private void Button_FillText(object sender, RoutedEventArgs e)
        {
            Excel.Worksheet ws = _xlApp.ActiveSheet;

            var array = new object[,] { { "apples", 5, 5 }, { "pears", 2, 2 }, { "oranges", 3, 3 }, { "apples", 5, 5 }, { "pears", 6, 6 } };
            Excel.Range Fruits = _xlApp.get_Range("E1", "G5");
            Excel.Range range = ws.Range[ws.Cells[1, 5], ws.Cells[1, 5]]; // row 1 col 1 to row 2 col 5
            ws.Names.Add("FruitRange", range);//"Sheet1!FruitRange"

            Fruits.Value = array;
            Fruits.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGreen);
        }
        private void Button_FindText(object sender, RoutedEventArgs e)
        {
            Excel.Range currentFind = null;
            Excel.Range firstFind = null;
            Excel.Worksheet ws = _xlApp.ActiveSheet;

            Excel.Range Fruits = _xlApp.get_Range("E1", "G5");
            // You should specify all these parameters every time you call this method,
            // since they can be overridden in the user interface. 
            //currentFind = Fruits.Find(What : "apples", After: null,
            //    LookIn:Excel.XlFindLookIn.xlValues, LookAt:Excel.XlLookAt.xlPart,
            //    SearchOrder : Excel.XlSearchOrder.xlByRows, SearchDirection:Excel.XlSearchDirection.xlNext, MatchCase:false,
            //    MatchByte:null, SearchFormat:null);
            //e Find(object What, object After, object LookIn, object LookAt, object SearchOrder, XlSearchDirection SearchDirection = XlSearchDirection.xlNext, 
            //    object MatchCase = null, object MatchByte = null, object SearchFormat = null);
            currentFind = Fruits.Find("apples", Type.Missing,
                   Excel.XlFindLookIn.xlValues, Excel.XlLookAt.xlPart,
                   Excel.XlSearchOrder.xlByRows, Excel.XlSearchDirection.xlNext, false,
                   Type.Missing, Type.Missing);

            while (currentFind != null)
            {

                // Keep track of the first range you find. 
                if (firstFind == null)
                {
                    firstFind = currentFind;
                }

                // If you didn't move to a new range, you are done.
                else if (currentFind.get_Address(Excel.XlReferenceStyle.xlA1)
                      == firstFind.get_Address(Excel.XlReferenceStyle.xlA1))
                {
                    break;
                }

                currentFind.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);
                currentFind.Font.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Red);
                currentFind.Font.Bold = true;

                //sheet.get_Range("7:9,12:12,14:14", Type.Missing) // range of rows
                //sheet.get_Range("7:9,12:12,14:14", Type.Missing) // range of rows
                var row = currentFind.Row;
                var col = currentFind.Column;
                Excel.Range Numbers = ws.Range[ws.Cells[row, col + 1], ws.Cells[row, col + 1]];
                // Excel.Range Numbers = ws.get_Range(ws.Cells[row, col+1], ws.Cells[row, col + 1]);
                Numbers.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Orange);

                currentFind = Fruits.FindNext(currentFind);
            }
        }
        private void Button_BorderFruitRange(object sender, RoutedEventArgs e)
        {
            Excel.Range range = getNamedRange("FruitRange");
            var c = range.Column;
            var r = range.Row;
            Excel.Worksheet ws = _xlApp.ActiveSheet;

            while (ws.Cells[r, c].Value != null)
            {
                c++;
            }
            c = c-1;
            while (ws.Cells[r, c].Value != null)
            {
                r++;
            }
            r = r - 1;
            var newrange = ws.Range[ws.Cells[range.Row, range.Column], ws.Cells[r, c]];
            newrange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.Yellow);

        }
        private Excel.Range getNamedRange(string name = "ChrisRange")
        {
            Excel.Worksheet ws = _xlApp.ActiveSheet;

            // find the range "Sheet1!ChrisRange" and format it. 
            foreach (Excel.Name n in ws.Names)
            {
                Console.WriteLine(n.Name);
                var chrisRangeName = $"{ws.Name}!{name}";
                if (n.Name == chrisRangeName)
                {
                    Excel.Range range = n.RefersToRange;
                    return range;
                }

            }

            return null;
        }
    }
}
/*
-		_xlApp	{System.__ComObject}	Microsoft.Office.Interop.Excel.Application {System.__ComObject}
		Native View	To inspect the native object, enable native code debugging.	
+		Non-Public members		
-		Dynamic View	Expanding the Dynamic View will get the dynamic members for the object	
		_Default	"Microsoft Excel"	System.String
+		ActiveCell	{System.__ComObject}	System.__ComObject
		ActiveChart	null	<null>
		ActiveDialog	null	<null>
		ActiveEncryptionSession	-1	System.Int32
+		ActiveMenuBar	{System.__ComObject}	System.__ComObject
		ActivePrinter	"Send To OneNote 2016 on nul:"	System.String
		ActiveProtectedViewWindow	null	<null>
+		ActiveSheet	{System.__ComObject}	System.__ComObject
+		ActiveWindow	{System.__ComObject}	System.__ComObject
+		ActiveWorkbook	{System.__ComObject}	System.__ComObject
+		AddIns	{System.__ComObject}	System.__ComObject
+		AddIns2	{System.__ComObject}	System.__ComObject
		AlertBeforeOverwriting	true	System.Boolean
		AltStartupPath	""	System.String
		AlwaysUseClearType	false	System.Boolean
		AnswerWizard	{System.Reflection.TargetInvocationException: Exception has been thrown by the target of an invocation. ---> System.Runtime.InteropServices.COMException: Exception from HRESULT: 0x800A03EC
   --- End of inner exception stack trace ---
   at System.RuntimeType.InvokeDispMethod(String name, BindingFlags invokeAttr, Object target, Object[] args, Boolean[] byrefModifiers, Int32 culture, String[] namedParameters)
   at System.RuntimeType.InvokeMember(String name, BindingFlags bindingFlags, Binder binder, Object target, Object[] providedArgs, ParameterModifier[] modifiers, CultureInfo culture, String[] namedParams)
   at System.Dynamic.IDispatchComObject.GetMembers(IEnumerable`1 names)}	System.Reflection.TargetInvocationException
+		Application	{System.__ComObject}	System.__ComObject
		ArbitraryXMLSupportAvailable	true	System.Boolean
		AskToUpdateLinks	true	System.Boolean
+		Assistance	{System.__ComObject}	System.__ComObject
+		Assistant	{System.__ComObject}	System.__ComObject
+		AutoCorrect	{System.__ComObject}	System.__ComObject
		AutoFormatAsYouTypeReplaceHyperlinks	true	System.Boolean
		AutomationSecurity	1	System.Int32
		AutoPercentEntry	true	System.Boolean
+		AutoRecover	{System.__ComObject}	System.__ComObject
		Build	6769	System.Double
		CalculateBeforeSave	true	System.Boolean
		Calculation	-4105	System.Int32
		CalculationInterruptKey	2	System.Int32
		CalculationState	0	System.Int32
		CalculationVersion	171027	System.Int32
		CanPlaySounds	true	System.Boolean
		CanRecordSounds	true	System.Boolean
		Caption	"Book2 - Excel"	System.String
		CellDragAndDrop	true	System.Boolean
+		Cells	{System.__ComObject}	System.__ComObject
		ChartDataPointTrack	true	System.Boolean
+		Charts	{System.__ComObject}	System.__ComObject
		ClusterConnector	""	System.String
		ColorButtons	true	System.Boolean
+		Columns	{System.__ComObject}	System.__ComObject
+		COMAddIns	{System.__ComObject}	System.__ComObject
+		CommandBars	{System.__ComObject}	System.__ComObject
		CommandUnderlines	-4105	System.Int32
		ConstrainNumeric	false	System.Boolean
		ControlCharacters	0	System.Double
		CopyObjectsWithCells	true	System.Boolean
		Creator	1480803660	System.Int32
		Cursor	-4143	System.Int32
		CursorMovement	1	System.Double
		CustomListCount	4	System.Double
		CutCopyMode	0	System.Int32
		DataEntryMode	-4146	System.Int32
		DDEAppReturnCode	0	System.Double
		DecimalSeparator	"."	System.String
		DefaultFilePath	"C:\\Users\\Waldo\\Documents"	System.String
		DefaultSaveFormat	51	System.Int32
		DefaultSheetDirection	-5003	System.Int32
+		DefaultWebOptions	{System.__ComObject}	System.__ComObject
		DeferAsyncQueries	false	System.Boolean
+		Dialogs	{System.__ComObject}	System.__ComObject
+		DialogSheets	{System.__ComObject}	System.__ComObject
		DisplayAlerts	true	System.Boolean
		DisplayClipboardWindow	false	System.Boolean
		DisplayCommentIndicator	-1	System.Int32
		DisplayDocumentActionTaskPane	false	System.Boolean
		DisplayDocumentInformationPanel	false	System.Boolean
		DisplayExcel4Menus	false	System.Boolean
		DisplayFormulaAutoComplete	true	System.Boolean
		DisplayFormulaBar	true	System.Boolean
		DisplayFullScreen	false	System.Boolean
		DisplayFunctionToolTips	true	System.Boolean
		DisplayInfoWindow	false	System.Boolean
		DisplayInsertOptions	true	System.Boolean
		DisplayNoteIndicator	true	System.Boolean
		DisplayPasteOptions	true	System.Boolean
		DisplayRecentFiles	true	System.Boolean
		DisplayScrollBars	true	System.Boolean
		DisplayStatusBar	true	System.Boolean
		Dummy101	{System.Reflection.TargetInvocationException: Exception has been thrown by the target of an invocation. ---> System.Runtime.InteropServices.COMException: Exception from HRESULT: 0x800A03EC
   --- End of inner exception stack trace ---
   at System.RuntimeType.InvokeDispMethod(String name, BindingFlags invokeAttr, Object target, Object[] args, Boolean[] byrefModifiers, Int32 culture, String[] namedParameters)
   at System.RuntimeType.InvokeMember(String name, BindingFlags bindingFlags, Binder binder, Object target, Object[] providedArgs, ParameterModifier[] modifiers, CultureInfo culture, String[] namedParams)
   at System.Dynamic.IDispatchComObject.GetMembers(IEnumerable`1 names)}	System.Reflection.TargetInvocationException
		Dummy22	{System.Reflection.TargetInvocationException: Exception has been thrown by the target of an invocation. ---> System.NotImplementedException: Not implemented (Exception from HRESULT: 0x80004001 (E_NOTIMPL))
   --- End of inner exception stack trace ---
   at System.RuntimeType.InvokeDispMethod(String name, BindingFlags invokeAttr, Object target, Object[] args, Boolean[] byrefModifiers, Int32 culture, String[] namedParameters)
   at System.RuntimeType.InvokeMember(String name, BindingFlags bindingFlags, Binder binder, Object target, Object[] providedArgs, ParameterModifier[] modifiers, CultureInfo culture, String[] namedParams)
   at System.Dynamic.IDispatchComObject.GetMembers(IEnumerable`1 names)}	System.Reflection.TargetInvocationException
		Dummy23	{System.Reflection.TargetInvocationException: Exception has been thrown by the target of an invocation. ---> System.NotImplementedException: Not implemented (Exception from HRESULT: 0x80004001 (E_NOTIMPL))
   --- End of inner exception stack trace ---
   at System.RuntimeType.InvokeDispMethod(String name, BindingFlags invokeAttr, Object target, Object[] args, Boolean[] byrefModifiers, Int32 culture, String[] namedParameters)
   at System.RuntimeType.InvokeMember(String name, BindingFlags bindingFlags, Binder binder, Object target, Object[] providedArgs, ParameterModifier[] modifiers, CultureInfo culture, String[] namedParams)
   at System.Dynamic.IDispatchComObject.GetMembers(IEnumerable`1 names)}	System.Reflection.TargetInvocationException
		EditDirectlyInCell	true	System.Boolean
		EnableAnimations	true	System.Boolean
		EnableAutoComplete	true	System.Boolean
		EnableCancelKey	1	System.Int32
		EnableCheckFileExtensions	true	System.Boolean
		EnableEvents	true	System.Boolean
		EnableLargeOperationAlert	true	System.Boolean
		EnableLivePreview	true	System.Boolean
		EnableMacroAnimations	false	System.Boolean
		EnableSound	false	System.Boolean
		EnableTipWizard	false	System.Boolean
+		ErrorCheckingOptions	{System.__ComObject}	System.__ComObject
+		Excel4IntlMacroSheets	{System.__ComObject}	System.__ComObject
+		Excel4MacroSheets	{System.__ComObject}	System.__ComObject
		ExtendList	true	System.Boolean
		FeatureInstall	0	System.Int32
+		FileExportConverters	{System.__ComObject}	System.__ComObject
		FileFind	{System.Reflection.TargetInvocationException: Exception has been thrown by the target of an invocation. ---> System.NotImplementedException: Not implemented (Exception from HRESULT: 0x80004001 (E_NOTIMPL))
   --- End of inner exception stack trace ---
   at System.RuntimeType.InvokeDispMethod(String name, BindingFlags invokeAttr, Object target, Object[] args, Boolean[] byrefModifiers, Int32 culture, String[] namedParameters)
   at System.RuntimeType.InvokeMember(String name, BindingFlags bindingFlags, Binder binder, Object target, Object[] providedArgs, ParameterModifier[] modifiers, CultureInfo culture, String[] namedParams)
   at System.Dynamic.IDispatchComObject.GetMembers(IEnumerable`1 names)}	System.Reflection.TargetInvocationException
		FileSearch	{System.Reflection.TargetInvocationException: Exception has been thrown by the target of an invocation. ---> System.NotImplementedException: Not implemented (Exception from HRESULT: 0x80004001 (E_NOTIMPL))
   --- End of inner exception stack trace ---
   at System.RuntimeType.InvokeDispMethod(String name, BindingFlags invokeAttr, Object target, Object[] args, Boolean[] byrefModifiers, Int32 culture, String[] namedParameters)
   at System.RuntimeType.InvokeMember(String name, BindingFlags bindingFlags, Binder binder, Object target, Object[] providedArgs, ParameterModifier[] modifiers, CultureInfo culture, String[] namedParams)
   at System.Dynamic.IDispatchComObject.GetMembers(IEnumerable`1 names)}	System.Reflection.TargetInvocationException
		FileValidation	0	System.Int32
		FileValidationPivot	0	System.Int32
+		FindFormat	{System.__ComObject}	System.__ComObject
		FixedDecimal	false	System.Boolean
		FixedDecimalPlaces	2	System.Int32
		FlashFill	true	System.Boolean
		FlashFillMode	false	System.Boolean
		FormulaBarHeight	1	System.Int32
		GenerateGetPivotData	false	System.Boolean
		GenerateTableRefs	1	System.Int32
		Height	558	System.Double
		HighQualityModeForGraphics	false	System.Boolean
		Hinstance	13631488	System.Int32
		HinstancePtr	13631488	System.Int32
		Hwnd	1249164	System.Int32
		IgnoreRemoteRequests	false	System.Boolean
		Interactive	true	System.Boolean
		IsSandboxed	false	System.Boolean
		Iteration	false	System.Boolean
+		LanguageSettings	{System.__ComObject}	System.__ComObject
		LargeButtons	false	System.Boolean
		LargeOperationCellThousandCount	33554	System.Int32
		Left	-5	System.Double
		LibraryPath	"C:\\Program Files (x86)\\Microsoft Office\\root\\Office16\\LIBRARY"	System.String
		MailSession	{}	System.DBNull
		MailSystem	1	System.Double
		MapPaperSize	true	System.Boolean
		MathCoprocessorAvailable	true	System.Boolean
		MaxChange	0.001	System.Double
		MaxIterations	100	System.Double
		MeasurementUnit	0	System.Int32
		MemoryFree	-2146826246	System.Int32
		MemoryTotal	-2146826246	System.Int32
		MemoryUsed	-2146826246	System.Int32
+		MenuBars	{System.__ComObject}	System.__ComObject
		MergeInstances	true	System.Boolean
+		Modules	{System.__ComObject}	System.__ComObject
		MouseAvailable	true	System.Boolean
		MoveAfterReturn	true	System.Boolean
		MoveAfterReturnDirection	-4121	System.Int32
+		MultiThreadedCalculation	{System.__ComObject}	System.__ComObject
		Name	"Microsoft Excel"	System.String
+		Names	{System.__ComObject}	System.__ComObject
		NetworkTemplatesPath	""	System.String
+		NewWorkbook	{System.__ComObject}	System.__ComObject
+		ODBCErrors	{System.__ComObject}	System.__ComObject
		ODBCTimeout	45	System.Int32
+		OLEDBErrors	{System.__ComObject}	System.__ComObject
		OnCalculate	null	<null>
		OnData	null	<null>
		OnDoubleClick	null	<null>
		OnEntry	null	<null>
		OnSheetActivate	null	<null>
		OnSheetDeactivate	null	<null>
		OnWindow	null	<null>
		OperatingSystem	"Windows (32-bit) NT 6.01"	System.String
		OrganizationName	""	System.String
+		Parent	{System.__ComObject}	System.__ComObject
		Path	"C:\\Program Files (x86)\\Microsoft Office\\root\\Office16"	System.String
		PathSeparator	"\\"	System.String
		PivotTableSelection	false	System.Boolean
		PrintCommunication	true	System.Boolean
		ProductCode	"{90160000-000F-0000-0000-0000000FF1CE}"	System.String
		PromptForSummaryInfo	false	System.Boolean
+		ProtectedViewWindows	{System.__ComObject}	System.__ComObject
+		QuickAnalysis	{System.__ComObject}	System.__ComObject
		Quitting	false	System.Boolean
		Ready	true	System.Boolean
+		RecentFiles	{System.__ComObject}	System.__ComObject
		RecordRelative	false	System.Boolean
		ReferenceStyle	1	System.Int32
+		ReplaceFormat	{System.__ComObject}	System.__ComObject
		RollZoom	false	System.Boolean
+		Rows	{System.__ComObject}	System.__ComObject
+		RTD	{System.__ComObject}	System.__ComObject
		SaveISO8601Dates	true	System.Boolean
		ScreenUpdating	true	System.Boolean
+		Selection	{System.__ComObject}	System.__ComObject
+		Sheets	{System.__ComObject}	System.__ComObject
		SheetsInNewWorkbook	1	System.Double
		ShowChartTipNames	true	System.Boolean
		ShowChartTipValues	true	System.Boolean
		ShowDevTools	false	System.Boolean
		ShowMenuFloaties	true	System.Boolean
		ShowQuickAnalysis	true	System.Boolean
		ShowSelectionFloaties	true	System.Boolean
		ShowStartupDialog	false	System.Boolean
		ShowToolTips	true	System.Boolean
		ShowWindowsInTaskbar	true	System.Boolean
+		SmartArtColors	{System.__ComObject}	System.__ComObject
+		SmartArtLayouts	{System.__ComObject}	System.__ComObject
+		SmartArtQuickStyles	{System.__ComObject}	System.__ComObject
+		SmartTagRecognizers	{System.__ComObject}	System.__ComObject
+		Speech	{System.__ComObject}	System.__ComObject
+		SpellingOptions	{System.__ComObject}	System.__ComObject
		StandardFont	"Calibri"	System.String
		StandardFontSize	11	System.Double
		StartupPath	"C:\\Users\\Waldo\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART"	System.String
		StatusBar	false	System.Boolean
		TemplatesPath	"C:\\Users\\Waldo\\AppData\\Roaming\\Microsoft\\Templates\\"	System.String
		ThisCell	{System.Reflection.TargetInvocationException: Exception has been thrown by the target of an invocation. ---> System.Runtime.InteropServices.COMException: Exception from HRESULT: 0x800A03EC
   --- End of inner exception stack trace ---
   at System.RuntimeType.InvokeDispMethod(String name, BindingFlags invokeAttr, Object target, Object[] args, Boolean[] byrefModifiers, Int32 culture, String[] namedParameters)
   at System.RuntimeType.InvokeMember(String name, BindingFlags bindingFlags, Binder binder, Object target, Object[] providedArgs, ParameterModifier[] modifiers, CultureInfo culture, String[] namedParams)
   at System.Dynamic.IDispatchComObject.GetMembers(IEnumerable`1 names)}	System.Reflection.TargetInvocationException
		ThisWorkbook	{System.Reflection.TargetInvocationException: Exception has been thrown by the target of an invocation. ---> System.Runtime.InteropServices.COMException: Exception from HRESULT: 0x800A03EC
   --- End of inner exception stack trace ---
   at System.RuntimeType.InvokeDispMethod(String name, BindingFlags invokeAttr, Object target, Object[] args, Boolean[] byrefModifiers, Int32 culture, String[] namedParameters)
   at System.RuntimeType.InvokeMember(String name, BindingFlags bindingFlags, Binder binder, Object target, Object[] providedArgs, ParameterModifier[] modifiers, CultureInfo culture, String[] namedParams)
   at System.Dynamic.IDispatchComObject.GetMembers(IEnumerable`1 names)}	System.Reflection.TargetInvocationException
		ThousandsSeparator	","	System.String
+		Toolbars	{System.__ComObject}	System.__ComObject
		Top	-6	System.Double
		TransitionMenuKey	"/"	System.String
		TransitionMenuKeyAction	1	System.Double
		TransitionNavigKeys	false	System.Boolean
		UILanguage	0	System.Int32
		UsableHeight	376.5	System.Double
		UsableWidth	1024.5	System.Double
		UseClusterConnector	false	System.Boolean
+		UsedObjects	{System.__ComObject}	System.__ComObject
		UserControl	true	System.Boolean
		UserLibraryPath	"C:\\Users\\Waldo\\AppData\\Roaming\\Microsoft\\AddIns\\"	System.String
		UserName	"Waldo"	System.String
		UseSystemSeparators	true	System.Boolean
		Value	"Microsoft Excel"	System.String
		VBE	{System.Reflection.TargetInvocationException: Exception has been thrown by the target of an invocation. ---> System.Runtime.InteropServices.COMException: Programmatic access to Visual Basic Project is not trusted

   --- End of inner exception stack trace ---
   at System.RuntimeType.InvokeDispMethod(String name, BindingFlags invokeAttr, Object target, Object[] args, Boolean[] byrefModifiers, Int32 culture, String[] namedParameters)
   at System.RuntimeType.InvokeMember(String name, BindingFlags bindingFlags, Binder binder, Object target, Object[] providedArgs, ParameterModifier[] modifiers, CultureInfo culture, String[] namedParams)
   at System.Dynamic.IDispatchComObject.GetMembers(IEnumerable`1 names)}	System.Reflection.TargetInvocationException
		Version	"16.0"	System.String
		Visible	true	System.Boolean
		WarnOnFunctionNameConflict	true	System.Boolean
+		Watches	{System.__ComObject}	System.__ComObject
		Width	1036.5	System.Double
+		Windows	{System.__ComObject}	System.__ComObject
		WindowsForPens	false	System.Boolean
		WindowState	-4137	System.Int32
+		Workbooks	{System.__ComObject}	System.__ComObject
+		WorksheetFunction	{System.__ComObject}	System.__ComObject
+		Worksheets	{System.__ComObject}	System.__ComObject




*/
