using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FREE_OSINT_Lib;
using Microsoft.Office.Interop.Word;
using System.Reflection;

namespace FREE_OSINT_Report_Builder
{
    public partial class Form1 : Form, IReport_module, IGeneral_module
    {
        private _Application oWord;
        private _Document oDoc;
        private object oRng;
        private object oEndOfDoc;
        private string description = "Generate word document using given treenodes.";
        private string title = "FREE-OSINT Report Builder";
        private WdBuiltinStyle[] preset_styles;

        public Form1()
        {
            preset_styles = new WdBuiltinStyle[5];
            preset_styles[0] = WdBuiltinStyle.wdStyleHeading1;
            preset_styles[1] = WdBuiltinStyle.wdStyleHeading2;
            preset_styles[2] = WdBuiltinStyle.wdStyleHeading3;
            preset_styles[3] = WdBuiltinStyle.wdStyleHeading4;
            preset_styles[4] = WdBuiltinStyle.wdStyleHeading5;
            //InitializeComponent();
            //GenerateDocument(new List<TreeNode>());
        }

        public string Description()
        {
            return description;
        }

        public object GenerateDocument(object infoToPdf)
        {
            List<TreeNode> treeNodes = (List<TreeNode>)infoToPdf;
            /*
            List<TreeNode> subnodes = new List<TreeNode>();
            subnodes.Add(new TreeNode("Vlad Adamko Profiles | Facebook"));
            subnodes.Add(new TreeNode("View the profiles of people named Vlad Adamko. Join Facebook to connect with Vlad Adamko and others you may know.Facebook gives people the power to..."));
            subnodes.Add(new TreeNode("https://www.facebook.com/public/Vlad-Adamko"));
            treeNodes.Add(new TreeNode("Vlad adamovych", subnodes.ToArray()));*/

            testWord(treeNodes);
            /*
            // Create a temporary file
            string filename = String.Format("{0}_tempfile.pdf", Guid.NewGuid().ToString("D").ToUpper());
            s_document = new PdfDocument();
            s_document.Info.Title = "Generated using FREE-OSINT_Report_Builder";
            s_document.Info.Author = "FREE-OSINT";
            s_document.Info.Subject = "";
            s_document.Info.Keywords = "Report, OSINT";

            s_document.AddPage();
            XGraphics gfx = XGraphics.FromPdfPage(s_document.Pages[0]);
            s_document = Utils.DrawTitle(s_document.Pages[0], gfx, "Report", s_document);

            // Save the s_document...
            s_document.Save(filename);
            // ...and start a viewer
            Process.Start(filename);*/
            return null;
        }
        public void testWord(List<TreeNode> treeNodes)
        {
            object oMissing = System.Reflection.Missing.Value;
            oEndOfDoc = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

            //Start Word and create a new document.

            oWord = new Microsoft.Office.Interop.Word.Application();
            oWord.Visible = true;
            oDoc = oWord.Documents.Add(ref oMissing, ref oMissing,
            ref oMissing, ref oMissing);

            //Insert a paragraph at the beginning of the document.
            Paragraph oPara1;
            oPara1 = oDoc.Content.Paragraphs.Add(ref oMissing);
            oPara1.Range.Font.Color = WdColor.wdColorBlack;
            oPara1.Range.Text = "Report";
            oPara1.Range.set_Style(WdBuiltinStyle.wdStyleTitle);
            oPara1.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
            oPara1.Range.Font.Name = "Times New Roman";
            oPara1.Range.Font.Bold = 1;
            oPara1.Format.SpaceAfter = 24;    //24 pt spacing after paragraph.
            oPara1.Range.InsertParagraphAfter();
            oPara1.Range.Font.Color = WdColor.wdColorBlack;


            //Insert a paragraph at the end of the document.
            foreach (TreeNode node in treeNodes)
            {
                Paragraph oPara2;
                oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                oPara2 = oDoc.Content.Paragraphs.Add(ref oRng);
                oPara2.Range.Text = node.Text;
                oPara2.Range.set_Style(preset_styles[0]);
                oPara2.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                oPara2.Range.Font.Name = "Times New Roman";
                oPara2.Range.Font.Bold = 1;
                oPara2.Format.SpaceAfter = 4;
                oPara2.Range.InsertParagraphAfter();
                if (node.Nodes.Count > 0)
                {
                    add_subnodes_to_paragraph(node, 1);
                }
            }

            /*


            //Insert a 3 x 5 table, fill it with data, and make the first row
            //bold and italic.
            Table oTable;
            Range wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 3, 5, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
            int r, c;
            string strText;
            for (r = 1; r <= 3; r++)
                for (c = 1; c <= 5; c++)
                {
                    strText = "r" + r + "c" + c;
                    oTable.Cell(r, c).Range.Text = strText;
                }
            oTable.Rows[1].Range.Font.Bold = 1;
            oTable.Rows[1].Range.Font.Italic = 1;

            //Add some text after the table.
            Paragraph oPara4;
            oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oPara4 = oDoc.Content.Paragraphs.Add(ref oRng);
            oPara4.Range.InsertParagraphBefore();
            oPara4.Range.Text = "And here's another table:";
            oPara4.Format.SpaceAfter = 24;
            oPara4.Range.InsertParagraphAfter();

            //Insert a 5 x 2 table, fill it with data, and change the column widths.
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oTable = oDoc.Tables.Add(wrdRng, 5, 2, ref oMissing, ref oMissing);
            oTable.Range.ParagraphFormat.SpaceAfter = 6;
            for (r = 1; r <= 5; r++)
                for (c = 1; c <= 2; c++)
                {
                    strText = "r" + r + "c" + c;
                    oTable.Cell(r, c).Range.Text = strText;
                }
            oTable.Columns[1].Width = oWord.InchesToPoints(2); //Change width of columns 1 & 2
            oTable.Columns[2].Width = oWord.InchesToPoints(3);

            //Keep inserting text. When you get to 7 inches from top of the
            //document, insert a hard page break.
            object oPos;
            double dPos = oWord.InchesToPoints(7);
            oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range.InsertParagraphAfter();
            do
            {
                wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                wrdRng.ParagraphFormat.SpaceAfter = 6;
                wrdRng.InsertAfter("A line of text");
                wrdRng.InsertParagraphAfter();
                oPos = wrdRng.get_Information
                                       (WdInformation.wdVerticalPositionRelativeToPage);
            }
            while (dPos >= Convert.ToDouble(oPos));
            object oCollapseEnd = WdCollapseDirection.wdCollapseEnd;
            object oPageBreak = WdBreakType.wdPageBreak;
            wrdRng.Collapse(ref oCollapseEnd);
            wrdRng.InsertBreak(ref oPageBreak);
            wrdRng.Collapse(ref oCollapseEnd);
            wrdRng.InsertAfter("We're now on page 2. Here's my chart:");
            wrdRng.InsertParagraphAfter();

            //Insert a chart.
            InlineShape oShape;
            object oClassType = "MSGraph.Chart.8";
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            oShape = wrdRng.InlineShapes.AddOLEObject(ref oClassType, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing,
            ref oMissing, ref oMissing, ref oMissing);

            //Demonstrate use of late bound oChart and oChartApp objects to
            //manipulate the chart object with MSGraph.
            object oChart;
            object oChartApp;
            oChart = oShape.OLEFormat.Object;
            oChartApp = oChart.GetType().InvokeMember("Application",
            BindingFlags.GetProperty, null, oChart, null);

            //Change the chart type to Line.
            object[] Parameters = new Object[1];
            Parameters[0] = 4; //xlLine = 4
            oChart.GetType().InvokeMember("ChartType", BindingFlags.SetProperty,
            null, oChart, Parameters);

            //Update the chart image and quit MSGraph.
            oChartApp.GetType().InvokeMember("Update",
            BindingFlags.InvokeMethod, null, oChartApp, null);
            oChartApp.GetType().InvokeMember("Quit",
            BindingFlags.InvokeMethod, null, oChartApp, null);
            //... If desired, you can proceed from here using the Microsoft Graph 
            //Object model on the oChart and oChartApp objects to make additional
            //changes to the chart.

            //Set the width of the chart.
            oShape.Width = oWord.InchesToPoints(6.25f);
            oShape.Height = oWord.InchesToPoints(3.57f);

            //Add text after the chart.
            wrdRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
            wrdRng.InsertParagraphAfter();
            wrdRng.InsertAfter("THE END.");
            */
        }

        public string Title()
        {
            return title;
        }

        private void add_subnodes_to_paragraph(TreeNode node, int level)
        {
            
            if (level == 1)
            {
                foreach (TreeNode treeNode in node.Nodes)
                {
                    //Insert another paragraph.
                    Paragraph oPara3;
                    oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    oPara3 = oDoc.Content.Paragraphs.Add(ref oRng);
                    oPara3.Range.Text = treeNode.Text;
                    oPara3.Range.set_Style(preset_styles[level]);
                    oPara3.Range.Font.Bold = 0;
                    oPara3.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    oPara3.Range.Font.Name = "Times New Roman";
                    oPara3.Format.SpaceAfter = 6;
                    oPara3.Range.InsertParagraphAfter();
                    if (treeNode.Nodes.Count > 0)
                    {
                        add_subnodes_to_paragraph(treeNode, level + 1);
                    }
                }
            }
            else
            {
                foreach (TreeNode treeNode in node.Nodes)
                {
                    //Insert another paragraph.
                    Paragraph oPara3;
                    oRng = oDoc.Bookmarks.get_Item(ref oEndOfDoc).Range;
                    oPara3 = oDoc.Content.Paragraphs.Add(ref oRng);
                    
                    oPara3.Range.Text = treeNode.Text;
                    if(level == 1)
                    {
                        oPara3.Range.set_Style(preset_styles[level]);
                    }
                    if (treeNode.Nodes.Count > 0)
                    {
                        oPara3.Range.set_Style(preset_styles[level]);
                    }
                    else
                    {
                        oPara3.Range.set_Style(WdBuiltinStyle.wdStyleBodyText);
                    }
                    oPara3.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphJustify;
                    oPara3.Range.Font.Name = "Times New Roman";
                    oPara3.Range.Font.Bold = 0;
                    oPara3.Format.SpaceAfter = 6;
                    oPara3.Range.InsertParagraphAfter();
                    if (treeNode.Nodes.Count > 0)
                    {
                        add_subnodes_to_paragraph(treeNode, level + 1);
                    }
                }
            }
        }
    }

}
