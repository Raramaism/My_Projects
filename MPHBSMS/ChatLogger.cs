using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Windows.Forms;

/// <summary>
/// Provides static methods for safely appending styled text to a RichTextBox.
/// </summary>
public static class ChatLogger
{
    // Method to append text with specific formatting
    public static void AppendText(RichTextBox rtb, string text, Color color, bool isBold = false)
    {
        // 1. Set insertion point to the end of the existing text
        rtb.SelectionStart = rtb.TextLength;
        rtb.SelectionLength = 0;

        // 2. Set the color and font style for the appended text
        rtb.SelectionColor = color;
        
        // This line requires the System.Drawing namespace (already included above)
        rtb.SelectionFont = new Font(rtb.Font, isBold ? FontStyle.Bold : FontStyle.Regular);
        
        // 3. Append the text followed by a new line
        rtb.AppendText(text + Environment.NewLine);
        
        // 4. Scroll to the end of the RichTextBox
        rtb.ScrollToCaret();
    }
}