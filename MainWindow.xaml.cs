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

		public MainWindow()
		{
			InitializeComponent();
		}

		private void SearchForContext()
		{
			var c = SearchTerms;
			ShowMessage($"Searching for {c}");
		}

		private void Embolden_Click(object sender, RoutedEventArgs e)
		{
			var sel = body_text.Selection;
		
			ShowMessage($"Bolding \"{sel.Text}\"");
			EditingCommands.ToggleBold.Execute(null, body_text);

			ShowMessage(SearchTerms);

		}

		//TODO: Break this out into a new object class
		private void MakeSlide()
		{
			ShowMessage("Opening Powerpoint...");

			var ppApp = new PowerPoint.Application();
			
			var pres = ppApp.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);

			var defaultSlide = pres.SlideMaster.CustomLayouts[3 /*PowerPoint.PpSlideLayout.ppLayoutTextAndObject*/ ];

			var sld = pres.Slides.AddSlide(1, defaultSlide);

			foreach (PowerPoint.Shape shp in sld.Shapes)
				shp.TextFrame.TextRange.Text = "Hello Powerpoint";

			pres.SaveAs("./new slide.pptx");

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

		private String SearchTerms{
			get
			{
				string terms = "";
				terms += title_text.Text.Trim();

				//find bold words
				var doc = body_text.Document;
				var txt = new TextRange(doc.ContentStart, doc.ContentEnd);
				var f = LogicalDirection.Forward;

				var tp = txt.Start.GetInsertionPosition(f);
				while (tp != null && tp.GetNextContextPosition(f) != null)
				{
					var phrase = new TextRange(tp, tp.GetNextContextPosition(f));

					var fw = phrase.GetPropertyValue(TextElement.FontWeightProperty);

					if (fw.Equals(FontWeights.Bold))
					{
						var boldterm = phrase.Text.Trim();
						if (boldterm.Length > 0)
							terms += " " + boldterm;
					}
					
					tp = tp.GetNextContextPosition(f);
				}

				//TODO: ADD BOLD WORDS TO SEARCH TERMS

				return terms;
			}
		}

		private void Button_Click(object sender, RoutedEventArgs e)
		{
			SearchForContext();
		}

		private String api_key = "AIzaSyABjUacUUZJHieTcqXM-k1teDLI-2oG0mk";

		private String google = "https://www.google.com/search?tbm=isch&q=shoebox";
		private string ddg = "https://duckduckgo.com/?q=shoebox&atb=v169-1&iax=images&ia=images";

		private void Body_text_TextChanged(object sender, TextChangedEventArgs e)
		{

		}
	}

}
