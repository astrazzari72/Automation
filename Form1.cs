using System;
using System.Diagnostics;
using System.Windows.Automation;

namespace Automation
{
	public partial class Form1 : Form
	{
		public Form1()
		{
			InitializeComponent();
		}

		private void button1_Click(object sender, EventArgs e)
		{
			string newText = textBox1.Text; //"\n(Appended) Scrivo altro testo in NOTEPAD da C#"; // text to append, starts with newline

			// Get the first Notepad process
			Process[] procs = Process.GetProcessesByName("notepad");
			if (procs.Length == 0)
			{
				Console.WriteLine("Notepad process not found.");
				return;
			}

			// Get AutomationElement from main window handle
			IntPtr hWnd = procs[0].MainWindowHandle;
			AutomationElement notepad = AutomationElement.FromHandle(hWnd);
			if (notepad == null)
			{
				Console.WriteLine("Notepad window not found.");
				return;
			}

			// Find the text area inside Notepad (Document OR Edit)
			var isEdit = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Edit);
			var isDoc = new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Document);
			var editOrDoc = new OrCondition(isEdit, isDoc);

			AutomationElement editor = notepad.FindFirst(TreeScope.Descendants, editOrDoc);
			if (editor == null)
			{
				Console.WriteLine("Text area (Edit/Document) not found.");
				return;
			}

			// Prefer ValuePattern for read/write; append to existing text
			if (editor.TryGetCurrentPattern(ValuePattern.Pattern, out object vpObj) && vpObj is ValuePattern vp)
			{
				string current = vp.Current.Value ?? string.Empty;

				// Ensure newline separation unless the current text already ends with newline
				bool endsWithNewline = current.EndsWith("\r\n") || current.EndsWith("\n") || current.Length == 0;
				string combined = current.Length == 0
				? newText
				: (endsWithWithNewline(current) ? current + newText : current + Environment.NewLine + newText);


				vp.SetValue(combined);
				Console.WriteLine("Text appended via ValuePattern.");
				return;
			}

			// Fallback: if ValuePattern not supported, try TextPattern (read) + notify user
			if (editor.TryGetCurrentPattern(TextPattern.Pattern, out object tpObj) && tpObj is TextPattern tp)
			{
				string current = tp.DocumentRange.GetText(-1) ?? string.Empty;
				Console.WriteLine("ValuePattern not supported; current text length: " + current.Length + ". Consider SendKeys fallback to append.");
				return;
			}

			Console.WriteLine("Neither ValuePattern nor TextPattern available. Cannot append.");
		}
		// Helper to check for newline ending robustly
		static bool endsWithWithNewline(string s) => s.EndsWith("\r\n") || s.EndsWith("\n");

		private void button2_Click(object sender, EventArgs e)
		{
			// Trova processo Notepad
			Process[] procs = Process.GetProcessesByName("notepad");
			if (procs.Length == 0) { Console.WriteLine("Notepad non avviato."); return; }

			var hWnd = procs[0].MainWindowHandle;
			var notepad = AutomationElement.FromHandle(hWnd);

			if (notepad == null) { Console.WriteLine("Finestra Notepad non trovata."); return; }

			// Trova menu "Visualizza"
			var viewMenu = notepad.FindFirst(TreeScope.Descendants,
				new AndCondition(
					new PropertyCondition(AutomationElement.NameProperty, "View"),
					new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)));

			if (viewMenu == null) { Console.WriteLine("Menu Visualizza non trovato."); return; }

			// Espandi il menu
			if (viewMenu.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out object expObj))
			{
				((ExpandCollapsePattern)expObj).Expand();
				Thread.Sleep(500);
			}

			// Dentro "Visualizza", trova "Zoom"
			var zoomMenu = viewMenu.FindFirst(TreeScope.Descendants,
				new AndCondition(
					new PropertyCondition(AutomationElement.NameProperty, "Zoom"),
					new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)));

			if (zoomMenu == null) { Console.WriteLine("Menu Zoom non trovato."); return; }

			if (zoomMenu.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out object zoomExp))
			{
				((ExpandCollapsePattern)zoomExp).Expand();
				Thread.Sleep(500);
			}

			// Dentro Zoom, trova “Aumenta”
			var zoomIn = zoomMenu.FindFirst(TreeScope.Descendants,
				new AndCondition(
					new PropertyCondition(AutomationElement.NameProperty, "Aumenta"),
					new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)));

			if (zoomIn == null) { Console.WriteLine("Voce Aumenta non trovata."); return; }

			// Esegui Invoke su “Aumenta”
			if (zoomIn.TryGetCurrentPattern(InvokePattern.Pattern, out object invObj))
			{
				((InvokePattern)invObj).Invoke();
				Console.WriteLine("Zoom + eseguito con successo.");
			}
		}
	}
}
