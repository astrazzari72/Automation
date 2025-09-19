using System;
using System.Diagnostics;
using System.Globalization;
using System.Threading;
using System.Windows.Automation;

public static class NotepadZoomHelper
{
	public static void OpenZoomInOnNotepad()
	{
		// Trova Notepad
		var procs = Process.GetProcessesByName("notepad");
		if (procs.Length == 0)
			throw new InvalidOperationException("Notepad non avviato.");

		var hWnd = procs[0].MainWindowHandle;
		var notepad = AutomationElement.FromHandle(hWnd);
		if (notepad == null)
			throw new InvalidOperationException("Finestra Notepad non trovata.");

		// Nomi possibili (EN/IT)
		string[] fileNames = { "File" };
		string[] newTabNames = { "New tab", "Nuova scheda" };

		// Condizioni
		Condition MenuItemNamed(params string[] names) =>
			new AndCondition(
				new OrCondition(Array.ConvertAll(names,
					n => (Condition)new PropertyCondition(AutomationElement.NameProperty, n))),
				new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.MenuItem)
			);

		// Trova File
		var fileMenu = WaitFind(notepad, MenuItemNamed(fileNames), TimeSpan.FromSeconds(3));
		if (fileMenu == null)
			throw new InvalidOperationException("Menu File non trovato.");

		TryExpand(fileMenu);

		// Trova New tab
		var newTab = WaitFind(fileMenu, MenuItemNamed(newTabNames), TimeSpan.FromSeconds(2));
		if (newTab == null)
			throw new InvalidOperationException("Voce 'New tab/Nuova scheda' non trovata.");

		// Invoca
		if (!TryInvoke(newTab))
			throw new InvalidOperationException("Impossibile invocare 'New tab/Nuova scheda'.");
	}

	// ====== Helper ======
	private static AutomationElement WaitFind(AutomationElement root, Condition cond, TimeSpan timeout)
	{
		var sw = Stopwatch.StartNew();
		AutomationElement elem = null;
		while (sw.Elapsed < timeout)
		{
			elem = root.FindFirst(TreeScope.Descendants, cond);
			if (elem != null) return elem;
			Thread.Sleep(100);
		}
		return null;
	}

	private static bool TryExpand(AutomationElement element)
	{
		if (element.TryGetCurrentPattern(ExpandCollapsePattern.Pattern, out object expObj))
		{
			var p = (ExpandCollapsePattern)expObj;
			if (p.Current.ExpandCollapseState != ExpandCollapseState.Expanded)
			{
				p.Expand();
				// attende che compaiano i figli
				Thread.Sleep(150);
			}
			return true;
		}

		// Alcuni menu reagiscono bene anche a Invoke come “apri”
		if (element.TryGetCurrentPattern(InvokePattern.Pattern, out object invObj))
		{
			((InvokePattern)invObj).Invoke();
			Thread.Sleep(150);
			return true;
		}

		return false;
	}

	private static bool TryInvoke(AutomationElement element)
	{
		if (element.TryGetCurrentPattern(InvokePattern.Pattern, out object invObj))
		{
			((InvokePattern)invObj).Invoke();
			return true;
		}

		// Fallback: prova LegacyIAccessible DefaultAction
		if (element.TryGetCurrentPattern(LegacyIAccessiblePattern.Pattern, out object accObj))
		{
			var acc = (LegacyIAccessiblePattern)accObj;
			try
			{
				acc.DoDefaultAction();
				return true;
			}
			catch { /* ignore */ }
		}

		return false;
	}
}
