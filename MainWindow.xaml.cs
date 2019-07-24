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
using System.Windows.Xps.Packaging;
using Path = System.IO.Path;

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

		public class ImageResult
		{
			public string title;
			public string imageUrl;
			public string thumbnailUrl;
			private BitmapImage thumb;
			private BitmapImage hires;

			public string dummy = "http://icons.iconarchive.com/icons/mattrich/adidas/512/Adidas-Shoebox-2-icon.png";

			public ImageResult(JObject item)
			{
				// json response -> "items"
				// items[i] -> title
				// items[i] -> link = img src
				// items[i] -> image -> thumbnailLink = thumbnail source

				if (item == null)
				{
					title = "Shoebox";
					imageUrl = dummy;
					thumbnailUrl = dummy;
				}
				else
				{
					title = item.Value<string>("title");
					imageUrl = item.Value<string>("link");
					var imageprops = item.Value<JObject>("image");
					thumbnailUrl = imageprops.Value<string>("thumbnailLink");
				}

				thumb = new BitmapImage();
				thumb.BeginInit();
				thumb.UriSource = new Uri(thumbnailUrl);
				thumb.CacheOption = BitmapCacheOption.OnLoad;
				thumb.EndInit();

			}

			public BitmapImage HiresImage {
				get {
					BitmapImage image = new BitmapImage();
					image.BeginInit();
					image.UriSource = new Uri(imageUrl);
					image.CacheOption = BitmapCacheOption.OnLoad;
					image.EndInit();
					return image;
				}
			}
			public BitmapImage ThumbnailImage {
				get {
					BitmapImage image = new BitmapImage();
					image.BeginInit();
					image.UriSource = new Uri(thumbnailUrl);
					image.CacheOption = BitmapCacheOption.OnLoad;
					image.EndInit();
					return image;
				}
			}


		}

		public List<ImageResult> Images { get; set; }

		private void SearchForContext()
		{
			if (listBox == null )
				return;

			String api_key = "AIzaSyABjUacUUZJHieTcqXM-k1teDLI-2oG0mk";
			String cx = "003899740029777113339:wt6hcdq1e04";
			String gsearch = $"https://www.googleapis.com/customsearch/v1?key={api_key}&cx={cx}&&searchType=image&q={SearchTerms}";

			String exurl = "https://www.googleapis.com/customsearch/v1?key=AIzaSyABjUacUUZJHieTcqXM-k1teDLI-2oG0mk&cx=003899740029777113339:wt6hcdq1e04&&searchType=image&q=shoebox";

			var gurl = new Uri(gsearch);
			ShowMessage($"Searching for {SearchTerms} at {gurl}");
			var req = PackWebRequest.CreateHttp(gurl);
			var res = req.GetResponse();

			var rsr = new StreamReader(res.GetResponseStream());
			string txt =  rsr.ReadToEnd() ?? "";

			var topj = JObject.Parse(txt);
			var items = topj["items"];

			Images = new List<ImageResult>();

			foreach (JObject item in items)
			{
				var ir = new ImageResult(item);
				Images.Add(ir);
			}

			//display all thumbnails
			listBox.Items.Clear();
			foreach (ImageResult ir in Images)
			{
				var lbi = new ListBoxItem();
				var image = new Image();
				image.Source = ir.ThumbnailImage;
				image.MaxHeight = listBox.ActualHeight;
				image.ToolTip = ir.title;
				lbi.Content = image;

				listBox.Items.Add(lbi);
			}
			ShowMessage("done");
		}


		//TODO: select up to 3 images

		private void Embolden_Click(object sender, RoutedEventArgs e)
		{
			var sel = body_text.Selection;
		
			ShowMessage($"Bolding \"{sel.Text}\"");
			EditingCommands.ToggleBold.Execute(null, body_text);

			ShowMessage(SearchTerms);


		}

		private void getDummyContent()
		{
			if (listBox == null)
				return;

			//make a bunch of thumbnails

			Images = new List<ImageResult>();
			while (Images.Count < 10)
				Images.Add(new ImageResult(null));

			listBox.Items.Clear();

			foreach (ImageResult ir in Images)
			{
				var lbi = new ListBoxItem();
				var image = new Image();
				image.Source = ir.ThumbnailImage;
				image.MaxHeight = listBox.ActualHeight;
				image.ToolTip = ir.title;
				lbi.Content = image;
				listBox.Items.Add(lbi);
			}

		}

		private void MakeSlide()
		{
			ShowMessage("Opening Powerpoint...");

			var pfile = new Microsoft.Win32.SaveFileDialog();
			pfile.DefaultExt = ".pptx";

			var doSave = pfile.ShowDialog() ?? false;
			if (!doSave)
				return;

			var ppApp = new PowerPoint.Application();
			var pres = ppApp.Presentations.Add(Microsoft.Office.Core.MsoTriState.msoTrue);
			var defaultSlide = pres.SlideMaster.CustomLayouts[4 /*PowerPoint.PpSlideLayout.ppLayoutTextAndObject*/ ];

			PowerPoint.Slide sld = pres.Slides.AddSlide(1, defaultSlide);
			//PowerPoint.Shape s = sld.Shapes[1];

			var sh_title = sld.Shapes[1];
			var sh_body = sld.Shapes[2];
			var sh_image = sld.Shapes[3];

			sh_title.TextFrame.TextRange.Text = title_text.Text;
			var doc = body_text.Document;
			sh_body.TextFrame.TextRange.Text = (new TextRange(doc.ContentStart, doc.ContentEnd)).Text;



			if (listBox.SelectedIndex != -1)
			{
				var img = Images.ElementAt<ImageResult>(listBox.SelectedIndex).HiresImage;
				sld.Shapes.AddPicture(img.UriSource.ToString(),
				Microsoft.Office.Core.MsoTriState.msoFalse, 
				Microsoft.Office.Core.MsoTriState.msoTrue,
				sh_image.Left, sh_image.Top, sh_image.Width);
			}

			//render preview

			/*
			//UPDATE PREVIEW
			String working_file = Path.GetTempPath() + "slide_builder.pptx";
			pres.SaveAs(working_file);

			string prev_file = Path.GetTempPath() + "slide_preview.xps";
			pres.ExportAsFixedFormat(prev_file, PowerPoint.PpFixedFormatType.ppFixedFormatTypeXPS);

			try
			{
				XpsDocument presxps = new XpsDocument(prev_file, FileAccess.Read);
				preview_box.Document = presxps.GetFixedDocumentSequence();
				preview_box.MaxWidth = preview_box.ViewportWidth;
			}
			catch (Exception e)
			{
				MessageBox.Show("couldn't open... " + e.ToString());
			}
			*/

			pres.SaveAs(pfile.FileName);
			//pres.Close();


		}

		private void SaveSlide(PowerPoint.Presentation pres)
		{

		}

		private void Make_slide_Click(object sender, RoutedEventArgs e)
		{	
			MakeSlide();
		}

		private void ShowMessage(String message)
		{
			if (message_block != null)
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
			getDummyContent();
			//SearchForContext();
		}

		private void Body_text_TextChanged(object sender, TextChangedEventArgs e)
		{
			
		}

		private void Title_text_TextChanged(object sender, TextChangedEventArgs e)
		{
			//would do this, but I would use up all of my API calls!

			//ShowMessage("refreshing results");
			//getDummyContent();
			//SearchForContext();
		}

		private void ListBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			if (listBox.SelectedIndex != -1)
			{
				var img = Images.ElementAt<ImageResult>(listBox.SelectedIndex).HiresImage;
				main_image.Source = img;
			}

		}
	}

}
