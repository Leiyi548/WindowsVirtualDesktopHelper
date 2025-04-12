using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace WindowsVirtualDesktopHelper {
	public partial class SwitchNotificationForm : Form {

		// P/Invoke declarations to hide the caret
		private const int WM_USER = 0x0400;
		private const int EM_HIDECARET = WM_USER + 21; // Edit control message to hide caret

		[DllImport("user32.dll", CharSet = CharSet.Auto)]
		private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

		public string LabelText;
		public string Position = "middlecenter";
		public int DisplayTimeMS = 2000;
		public int? ScreenNumber = null;
		public bool FadeIn = false;
		public bool Translucent = true;

		private int animationDirection = +1;
		private int animationOpacityCurrent = 0;
		private int animationOpacityTarget = 100;
		private int animationOpacityStep = 5;

		public static event EventHandler WillShowNotificationFormEvent;

		public SwitchNotificationForm(int? screenNumber = null) {
			InitializeComponent();

			this.ScreenNumber = screenNumber;
			this.FadeIn = Settings.GetBool("feature.showDesktopSwitchOverlay.animate");
			this.Translucent = Settings.GetBool("feature.showDesktopSwitchOverlay.translucent");
			this.Position = Settings.GetString("feature.showDesktopSwitchOverlay.position");

			// Theme
			var theme = App.Instance.CurrentSystemThemeName;
			var fgColor = ColorTranslator.FromHtml(Settings.GetString("theme.overlay.overlayFG." + theme));
			var bgColor = ColorTranslator.FromHtml(Settings.GetString("theme.overlay.overlayBG." + theme));
			this.BackColor = bgColor;
			this.richTextBox1.ForeColor = fgColor;
			this.richTextBox1.BackColor = bgColor;
			
			// Size
			this.Width = Settings.GetInt("theme.overlay.width");
			this.Height = Settings.GetInt("theme.overlay.height");

			// Font
			var font = new Font(Settings.GetString("theme.overlay.font"), Settings.GetFloat("theme.overlay.fontSize"), FontStyle.Regular);
			this.richTextBox1.Font = font;

			// Set position
			var positionOffset = 40;
			this.StartPosition = FormStartPosition.Manual;
			var screen = Screen.FromControl(this); // get main screen
			if(this.ScreenNumber != null) screen = Screen.AllScreens[this.ScreenNumber.Value]; 
			var screenW = screen.WorkingArea.Width;
			var screenH = screen.WorkingArea.Height;
			var screenX = screen.WorkingArea.X;
			var screenY = screen.WorkingArea.Y;
			if (this.Position == "topleft") {
				this.Location = new Point(screenX+positionOffset, screenY + positionOffset);
			} else if (this.Position == "topcenter") {
				this.Location = new Point(screenX + screenW / 2 - this.Width / 2, screenY + positionOffset);
			} else if (this.Position == "topright") {
				this.Location = new Point(screenX + screenW - this.Width - positionOffset, screenY + positionOffset);
			} else if (this.Position == "middleleft") {
				this.Location = new Point(screenX + positionOffset, screenY + screenH / 2 - this.Height / 2);
			} else if (this.Position == "middlecenter") {
				this.Location = new Point(screenX + screenW / 2 - this.Width / 2, screenY + screenH / 2 - this.Height / 2);
			} else if (this.Position == "middleright") {
				this.Location = new Point(screenX + screenW - this.Width - positionOffset, screenY + screenH / 2 - this.Height / 2);
			} else if (this.Position == "bottomleft") {
				this.Location = new Point(screenX + positionOffset, screenY + screenH - this.Height - positionOffset);
			} else if (this.Position == "bottomcenter") {
				this.Location = new Point(screenX + screenW / 2 - this.Width / 2, screenY + screenH - this.Height - positionOffset);
			} else if (this.Position == "bottomright") {
				this.Location = new Point(screenX + screenW - this.Width - positionOffset, screenY + screenH - this.Height - positionOffset);
			}

			
			// Register for signal to close self
			WillShowNotificationFormEvent += this.OnWillShowNotificationForm;

			// Set initial opacity
			if (this.FadeIn == true) {
				this.Opacity = 0;
			} else {
				if (this.Translucent) this.Opacity = 0.6;
				else this.Opacity = 1.0;
			}
		}

		public static void CloseAllNotifications(object sender) {
			// Send signal to close all others
			SwitchNotificationForm.WillShowNotificationFormEvent?.Invoke(sender, EventArgs.Empty);
		}

		protected virtual void OnWillShowNotificationForm(object sender, EventArgs e) {
			if (sender == this) return; // make sure we dont close ourselves

			// If another SwitchNotification is being shown, then close ourselves
			this.Close();
		}

		// Preven from from being focusable
		// via https://stackoverflow.com/questions/2423234/make-a-form-not-focusable-in-c-sharp
		// via https://stackoverflow.com/questions/2798245/click-through-in-c-sharp-form
		private const int WS_EX_NOACTIVATE = 0x08000000;
		private const int WS_EX_LAYERED = 0x80000;
		// private const int WS_EX_TRANSPARENT = 0x20; // Temporarily disable for RichTextBox testing
		protected override CreateParams CreateParams {
			get {
				var createParams = base.CreateParams;

				createParams.ExStyle |= WS_EX_NOACTIVATE;
				createParams.ExStyle |= WS_EX_LAYERED;
				// createParams.ExStyle |= WS_EX_TRANSPARENT; // Temporarily disable
				return createParams;
			}
		}

		private void SwitchNotificationForm_Shown(object sender, EventArgs e) {
			this.timerClose.Interval = this.DisplayTimeMS;

			if (this.FadeIn) {
				this.animationOpacityCurrent = 0;
				this.animationOpacityTarget = 100;
				if (this.Translucent) this.animationOpacityTarget = 60;
				else this.animationOpacityTarget = 100;
				this.timerAnimate.Start();
			} else {
				this.timerClose.Start();
			}
			// Also hide caret when shown, just in case
			SendMessage(this.richTextBox1.Handle, EM_HIDECARET, IntPtr.Zero, IntPtr.Zero);
		}

		private void SwitchNotificationForm_Load(object sender, EventArgs e) {
			// Text setting and coloring is now handled by ApplyTextFormatting called from App.cs
			// Center align text initially (optional, can be done in ApplyTextFormatting too)
			this.richTextBox1.SelectAll();
			this.richTextBox1.SelectionAlignment = HorizontalAlignment.Center;
			this.richTextBox1.DeselectAll();
		}

		private void timerClose_Tick(object sender, EventArgs e) {
			if(this.DisplayTimeMS == 0) return; // we dont close if DisplayTimeMS is 0

			if (this.FadeIn) {
				this.animationOpacityCurrent = this.animationOpacityTarget;
				this.animationOpacityTarget = 0;
				this.animationDirection = -1;
				this.timerAnimate.Start();
			} else {
				this.Close();
			}
			this.Opacity = this.animationOpacityCurrent / 100.0;
		}

		private void timerAnimate_Tick(object sender, EventArgs e) {
			this.animationOpacityCurrent += (this.animationOpacityStep * this.animationDirection);
			if (
				(this.animationDirection == +1 && this.animationOpacityCurrent >= this.animationOpacityTarget)
				||
				(this.animationDirection == -1 && this.animationOpacityCurrent <= this.animationOpacityTarget)
				) {
				this.animationOpacityCurrent = this.animationOpacityTarget;
				this.timerAnimate.Enabled = false;
				if (this.animationDirection == +1) {
					this.timerClose.Start();
				} else {
					this.Close();
				}
			}
			this.Opacity = this.animationOpacityCurrent / 100.0;
		}

		// Renamed and made public to be called from App.cs after LabelText is set
		public void ApplyTextFormatting(string fullText, string displayName, Color displayNameColor, string examName, Color examNameColor, string daysString, Color daysColor)
		{
			if (string.IsNullOrEmpty(fullText))
			{
				this.richTextBox1.Text = "";
				return;
			}

			// Get font info and colors
			var defaultFontName = this.richTextBox1.Font.Name;
			int rtfFontSize = (int)(this.richTextBox1.Font.SizeInPoints * 2);
			var defaultColor = this.richTextBox1.ForeColor;
			var bgColor = this.richTextBox1.BackColor;

			// Find indices and lengths (handle null/empty strings)
			int displayNameLength = displayName?.Length ?? 0;
			int examNameIndex = (!string.IsNullOrEmpty(examName) && !string.IsNullOrEmpty(fullText)) ? fullText.IndexOf(examName) : -1;
			int examNameLength = (examNameIndex != -1) ? examName.Length : 0;
			int daysIndex = (!string.IsNullOrEmpty(daysString) && !string.IsNullOrEmpty(fullText)) ? fullText.IndexOf(daysString, (examNameIndex != -1 ? examNameIndex + examNameLength : displayNameLength)) : -1; // Search after exam name or display name
			int daysLength = (daysIndex != -1) ? daysString.Length : 0;

			// Build RTF String
			var rtfBuilder = new System.Text.StringBuilder();
			rtfBuilder.Append(@"{\rtf1\ansi\deff0");
			rtfBuilder.Append($@"{{\fonttbl{{\f0 {defaultFontName};}}}}");
			// Color table: 1=DefaultFG, 2=DisplayNameColor, 3=ExamNameColor, 4=DaysColor
			rtfBuilder.Append($@"{{\colortbl ;{ColorToRtfFormat(defaultColor)};{ColorToRtfFormat(displayNameColor)};{ColorToRtfFormat(examNameColor)};{ColorToRtfFormat(daysColor)};}}");
			// Default paragraph format: centered, font 0, size
			rtfBuilder.Append($@"\pard\qc\f0\fs{rtfFontSize} ");

			// Iterate through the full text and apply colors segment by segment
			int currentPos = 0;
			while (currentPos < fullText.Length)
			{
				int nextSegmentStart = fullText.Length; // Default to end
				int segmentLength = 0;
				int colorIndex = 1; // Default color
				string segmentText = "";

				// Determine the next segment and its color
				if (currentPos == 0 && displayNameLength > 0) // Display Name segment
				{
					segmentLength = displayNameLength;
					colorIndex = 2;
				}
				else if (examNameIndex != -1 && currentPos == examNameIndex) // Exam Name segment
				{
					segmentLength = examNameLength;
					colorIndex = 3;
				}
				else if (daysIndex != -1 && currentPos == daysIndex) // Days segment
				{
					segmentLength = daysLength;
					colorIndex = 4;
				}
				else // Default color segment
				{
					// Find the beginning of the *next* colored segment
					int nextSpecialStart = fullText.Length;
					if (examNameIndex != -1 && examNameIndex > currentPos)
						nextSpecialStart = Math.Min(nextSpecialStart, examNameIndex);
					if (daysIndex != -1 && daysIndex > currentPos)
						nextSpecialStart = Math.Min(nextSpecialStart, daysIndex);
					
					segmentLength = nextSpecialStart - currentPos;
					colorIndex = 1;
				}

				// Ensure segment length is valid
				segmentLength = Math.Max(0, Math.Min(segmentLength, fullText.Length - currentPos));
				if (segmentLength == 0) break; // Prevent infinite loop if logic fails

				segmentText = fullText.Substring(currentPos, segmentLength);

				// Append the RTF for this segment
				rtfBuilder.Append($@"{{\cf{colorIndex} "); // Start color scope
				// Apply bold only if it's the days segment (colorIndex == 4)
				if (colorIndex == 4)
				{
					rtfBuilder.Append(@"\b "); // Turn bold ON
				}
				else
				{
					rtfBuilder.Append(@"\b0 "); // Ensure bold is OFF for other segments
				}
				// Append text with Unicode/RTF escaping
				foreach (char c in segmentText)
				{
					if (c == '\\') rtfBuilder.Append("\\\\");
					else if (c == '{') rtfBuilder.Append("\\{");
					else if (c == '}') rtfBuilder.Append("\\}");
					else if (c == '\n') rtfBuilder.Append("\\par ");
					else if (c <= 0x7f) rtfBuilder.Append(c);
					else rtfBuilder.Append($"\\u{(int)c}?");
				}
				rtfBuilder.Append(@"} "); // End color/bold scope

				// Move to the next position
				currentPos += segmentLength;
			}

			// End RTF document
			rtfBuilder.Append(@"}");

			try
			{
				this.richTextBox1.Rtf = rtfBuilder.ToString();
			}
			catch (Exception ex)
			{
				Util.Logging.WriteLine($"Error applying RTF with multiple colors: {ex.Message}. Fallback to plain text.");
				this.richTextBox1.Text = fullText; // Fallback
			}

			 // Re-apply properties
			this.richTextBox1.BackColor = bgColor;
			this.richTextBox1.ReadOnly = true;
			this.richTextBox1.BorderStyle = BorderStyle.None;
			this.richTextBox1.Cursor = Cursors.Default;

			// Deselect all
			this.richTextBox1.Select(0, 0);
			this.richTextBox1.DeselectAll();

			// Hide the caret using SendMessage
			SendMessage(this.richTextBox1.Handle, EM_HIDECARET, IntPtr.Zero, IntPtr.Zero);
		}

		private string ColorToRtfFormat(Color c)
		{
			return $"\\red{c.R}\\green{c.G}\\blue{c.B}";
		}
	}
}
