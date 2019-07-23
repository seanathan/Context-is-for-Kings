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
using System.IO.Packaging;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace Context_is_for_Kings
{
	/// <summary>
	/// Interaction logic for MainWindow.xaml
	/// </summary>
	public partial class MainWindow : Window
	{

		private String searchString = "";

		public MainWindow()
		{
			InitializeComponent();
		}

		private void Embolden_Click(object sender, RoutedEventArgs e)
		{
			var sel = body_text.Selection;
		
			if (sel == null ||
				!EditingCommands.ToggleBold.CanExecute(null, body_text))
				ShowMessage("nothing to bold");
			else{
				ShowMessage($"Bolding \"{sel.Text}\"");
				EditingCommands.ToggleBold.Execute(null, body_text);
			}

		}


		//TODO: Break this out into a new object class
		private void MakeSlide()
		{
			ShowMessage("Opening Powerpoint...");

			var ppApp = new PowerPoint.Application();
			ppApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;



		}

		private void Make_slide_Click(object sender, RoutedEventArgs e)
		{	
			MakeSlide();
		}

		private void ShowMessage(String message)
		{
			message_block.Text = message;
		}
	}
}
