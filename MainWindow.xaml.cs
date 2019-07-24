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
using System.Net;
using System.IO;
using NJS = Newtonsoft.Json;
using System.Web.Script.Serialization;
using Newtonsoft.Json.Linq;

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

		class ImageResult
		{
			public string title { get; set; }
			public string link { get; set; }
			public string thumbnail { get; set; }

	}

		private void SearchForContext()
		{
			String api_key = "AIzaSyABjUacUUZJHieTcqXM-k1teDLI-2oG0mk";
			String cx = "003899740029777113339:wt6hcdq1e04";
			String gsearch = $"https://www.googleapis.com/customsearch/v1?key={api_key}&cx={cx}&&searchType=image&q={SearchTerms}";

			String exurl = "https://www.googleapis.com/customsearch/v1?key=AIzaSyABjUacUUZJHieTcqXM-k1teDLI-2oG0mk&cx=003899740029777113339:wt6hcdq1e04&&searchType=image&q=shoebox";

			var gurl = new Uri(gsearch);

			ShowMessage($"Searching for {SearchTerms} at {gurl}");

			var req = PackWebRequest.CreateHttp(gurl);

			var res = req.GetResponse();


			debug_output.Text = "Results:\n";

			var rsr = new StreamReader(res.GetResponseStream());

			string txt =  rsr.ReadToEnd() ?? "";

			var topj = JObject.Parse(txt);
			var items = topj["items"];
			foreach (JObject item in items)
			{
				var ir = new ImageResult();
				ir.title = item.Value<string>("title");
				ir.link = item.Value<string>("link");
				ir.thumbnail = item.Value<string>("thumbnail");

				debug_output.Text += $"\n{ir.title}\n{ir.link}\n{ir.thumbnail}";
			}

			// json response -> "items"
			// items[i] -> title
			// items[i] -> link = img src
			// items[i] -> image -> thumbnailLink = thumbnail source

			ShowMessage("content type: " + res.ContentType);
		}


		//TODO: select up to 3 images

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

			var defaultSlide = pres.SlideMaster.CustomLayouts[4 /*PowerPoint.PpSlideLayout.ppLayoutTextAndObject*/ ];

			var sld = pres.Slides.AddSlide(1, defaultSlide);

			foreach (PowerPoint.Shape shp in sld.Shapes)
				shp.TextFrame.TextRange.Text = "Hello Powerpoint";

			pres.SaveAs("./new slide.pptx"); // need generated filenames

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

		private void Body_text_TextChanged(object sender, TextChangedEventArgs e)
		{
			
		}
	}

}
