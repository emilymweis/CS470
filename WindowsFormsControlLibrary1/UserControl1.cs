using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Documents;
using Microsoft.Office.Interop.Word;
using Table = Microsoft.Office.Interop.Word.Table;
using System.Collections;

namespace WindowsFormsControlLibrary1
{
    public partial class UserControl1: UserControl
    {
        public UserControl1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            createDocument();
        }

       public string returnOne()
        {
            string str = "";
            str = "Name of person" + Environment.NewLine +
            "Via: source of communication";


            return str;
        }

        public string returnTwo()
        {
            string str = "";
            str = Environment.NewLine+"Re:  Calculation of Value of ABC Company";


            return str;
        }

        public string returnThree()
        {
            string str = "";
            str = Environment.NewLine + "Dear Mr.Smith:" + Environment.NewLine +

            "Please find enclosed our calculation of the fair market value of a 100 % ownership interest (“Subject Interest”)" +
             " in ABC Company (“ABC” or the “Company”) as of December 31, 2018.The purpose of this report is to calculate the fair market value the Company for corporate planning purposes. Should additional information become available, we reserve the right to update our valuation analysis and report.";

            return str;
        }

        public string returnFour()
        {
            string str = "";
            str = "Based on the review of data and information provided, we find a calculated value of the Subject Interest in The Company on a controlling, marketable basis as of December 31, 2018 to be:";
            return str;
        }

        public string returnFifth()
        {
            string str = "";
            str = "We have performed a calculation engagement in accordance with the “Statement on Standards for Valuation Services No. 1” (SSVS) of the American Institute of Certified Public Accountants (AICPA). SSVS defines a calculation engagement as:";
            return str;
        }

        public string returnSix()
        {
            string str = "";
            str = "$13,658,000(change to new value)";
            return str;
        }

        public string returnSeven()
        {
            string str = "";
            str = "“An engagement to estimate value wherein the valuation analyst and the client agree on the specific valuation approaches and valuation methods that the valuation analyst will use and the extent of valuation procedures the valuation analyst will perform to estimate the value of a subject interest.  A calculation engagement generally does not include all of the valuation procedures required for a valuation engagement.  If a valuation engagement had been performed, the results might have been different.  The valuation analyst expresses the results of the calculation engagement as a calculated value, which may be either a single amount or a range.”";
            return str;
         }

        private void createDocument()
        {
            try
            {
                //Create an instance for word app  
                Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();

                //Set animation status for word application  
                winword.ShowAnimation = false;

                //Set status for word application is to be visible or not.  
                winword.Visible = false;

                //Create a missing variable for missing value  
                object missing = System.Reflection.Missing.Value;

                //Create a new document  
                Microsoft.Office.Interop.Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);


                Microsoft.Office.Interop.Word.Paragraph para4 = document.Content.Paragraphs.Add(ref missing);
                para4.Range.Font.Bold = 1;
                para4.Range.Font.Name = "Helvetica";
                //para4.Range.ParagraphFormat.IndentFirstLineCharWidth(34);
                //para4.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para4.Range.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle;
                para4.Range.Font.Size = 10;
                //para4.Range.ParagraphFormat.FirstLineIndent = 20;
                para4.LineSpacing = 11F;
                para4.Range.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlack;
                //para4.Range.ParagraphFormat.FirstLineIndent = 20;
               // para4.Range.Text = returnString();
                para4.Range.ParagraphFormat.TabIndent(20);
                para4.Range.InsertParagraphAfter();


                Microsoft.Office.Interop.Word.Paragraph para5 = document.Content.Paragraphs.Add(ref missing);
                
                //para4.Range.ParagraphFormat.IndentFirstLineCharWidth(34);
                //para4.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para5.Range.Font.Size = 11;
                //para4.Range.ParagraphFormat.FirstLineIndent = 20;
                para5.LineSpacing = 11F;
                para5.Range.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlack;
                //para4.Range.ParagraphFormat.FirstLineIndent = 20;
                para5.Range.Text = returnOne();
                para5.Range.Font.Name = "Helvetica";

                para5.Range.InsertParagraphAfter();


                Microsoft.Office.Interop.Word.Paragraph para6 = document.Content.Paragraphs.Add(ref missing);
                para6.Range.Font.Bold = 1;
                para6.Range.Font.Size = 11;
                //para4.Range.ParagraphFormat.FirstLineIndent = 20;
                para6.LineSpacing = 11F;
                para6.Range.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlack;
                //para4.Range.ParagraphFormat.FirstLineIndent = 20;
                para6.Range.Text = returnTwo();
                para6.Range.Font.Name = "Helvetica";
                para6.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para6.Range.InsertParagraphAfter();

               
                para6.Range.Font.Size = 11;
                //para4.Range.ParagraphFormat.FirstLineIndent = 20;
                para6.LineSpacing = 11F;
                para6.Range.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlack;
                //para4.Range.ParagraphFormat.FirstLineIndent = 20;
                para6.Range.Text = returnThree();
                para6.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para6.Range.InsertParagraphAfter();

               
                para6.Range.Bold = 1;
                para6.Range.Font.Size = 11;
                //para4.Range.ParagraphFormat.FirstLineIndent = 20;
                para6.LineSpacing = 11F;
                para6.Range.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlack;
                //para4.Range.ParagraphFormat.FirstLineIndent = 20;
                para6.Range.Text = returnFour();
                para6.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para6.Range.InsertParagraphAfter();


                para6.Range.Bold = 1;
                para6.Range.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle;
                para6.Range.Font.Size = 11;//needs to be underlined
                //para6.Range.ParagraphFormat.FirstLineIndent = 100; //this actually worked
                para6.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para6.LineSpacing = 11F;
                para6.Range.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlack;
                //para4.Range.ParagraphFormat.FirstLineIndent = 20;
                para6.Range.Text = returnSix();
                para6.Range.InsertParagraphAfter();
                para6.Range.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;//got to remove the underline format so it doesnt affect everything

                /* object numTimes3 = 7;
                 document.Undo(ref numTimes3);*/
                Microsoft.Office.Interop.Word.Paragraph para7 = document.Content.Paragraphs.Add(ref missing);
                para7.Range.Bold = 0;
                para7.Range.Font.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para7.Range.Font.Size = 11;
                para7.LineSpacing = 11F;
                para7.Range.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlack;
                para7.Range.Text = returnFifth();
                para7.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para7.Range.InsertParagraphAfter();

                Microsoft.Office.Interop.Word.Paragraph para8 = document.Content.Paragraphs.Add(ref missing);
                para8.Range.Bold = 0;
                para8.Range.Font.Size = 11;
                para8.LineSpacing = 11F;
                para8.Range.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlack;
                para8.Range.Text = returnSeven();
                para8.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para8.Range.ParagraphFormat.RightIndent = 35;
                para8.Range.ParagraphFormat.LeftIndent = 35;//these tabs are actually .5 on the tab margin word converts these number werid
                para8.Range.InsertParagraphAfter();


                ///keep in mind the order of how you try to format the paragraphs is very important
                ///similarly this code needs to be refactored but for now it works
                Microsoft.Office.Interop.Word.Paragraph para9 = document.Content.Paragraphs.Add(ref missing);
                para9.Range.Bold = 0;
                para9.Range.Font.Size = 11;
                para9.LineSpacing = 11F;
                para9.Range.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlack;
                para9.Range.Text = "SSVS addresses a calculation report as follows:";
                para9.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para8.Range.ParagraphFormat.RightIndent = 0;
                para9.Range.ParagraphFormat.LeftIndent = 0;
                para9.Range.InsertParagraphAfter();


                para8.Range.Bold = 0;
                para8.Range.Text = returnEight();
                para8.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para8.Range.ParagraphFormat.RightIndent = 35;
                para8.Range.ParagraphFormat.LeftIndent = 35;//these tabs are actually .5 on the tab margin word converts these number werid
                para8.Range.InsertParagraphAfter();

                Microsoft.Office.Interop.Word.Paragraph para10 = document.Content.Paragraphs.Add(ref missing);
                
                para10.Range.Bold = 0;
                para10.Range.Font.Size = 11;
                para10.LineSpacing = 11F;
                para10.Range.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlack;
                para10.Range.Text = returnNine();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Text = returnTen();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                Microsoft.Office.Interop.Word.Section wordSection = document.Content.Sections.Add(ref missing);
                //Get the footer range and add the footer details.  
                Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlack;
                footerRange.Font.Size = 10;
                footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                footerRange.Text = "1NACVA (National Association of Certified Valuators and Analysts)";

                //document.Words.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);//this should insert a page break. works
                //inserting a footer automatically inserts a page break??????? not sure

                para10.Range.Text = returnEleven();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para8.Range.Text = returnTwelve();
                para8.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para8.Range.ParagraphFormat.RightIndent = 35;
                para8.Range.ParagraphFormat.LeftIndent = 35;//these tabs are actually .5 on the tab margin word converts these number werid
                para8.Range.InsertParagraphAfter();

                para10.Range.Text = returnThirteen();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                ///para8 has the werid formetting indents
                para8.Range.Text = returnFourteen();
                para8.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para8.Range.ParagraphFormat.RightIndent = 35;
                para8.Range.ParagraphFormat.LeftIndent = 35;//these tabs are actually .5 on the tab margin word converts these number werid
                para8.Range.InsertParagraphAfter();

                para10.Range.Text = returnFifthteen();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para8.Range.Text = returnSixthteen();
                para8.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para8.Range.ParagraphFormat.RightIndent = 35;
                para8.Range.ParagraphFormat.LeftIndent = 35;
                para8.Range.InsertParagraphAfter();

                para10.Range.Text = returnSeventeen();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Text = returnEighteen();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Text = returnEighteen();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();


                //this whole assets thing is an attempt to make a bulleted list.
                Microsoft.Office.Interop.Word.Paragraph assets;
                assets = document.Content.Paragraphs.Add(Type.Missing);
                
                // Some code to generate the text
                ArrayList assetsList = new ArrayList();
                assetsList.Add("• The history and nature of the business");
                assetsList.Add("• The economic outlook of the United States and that of the specific industry in particular");
                assetsList.Add("• The book value of the subject company’s stock and the financial condition of the business");
                assetsList.Add("• The earning capacity of the company");
                assetsList.Add("• The dividend-paying capacity of the company");
                assetsList.Add("• Whether or not the company has goodwill or other intangible value");
                assetsList.Add("• Sales of the stock and size of the block of stock to be value");
                assetsList.Add("• The market price of publicly traded stocks of corporations engaged in similar industries or lines of business");

                string assetText = "";
                foreach (String asset in assetsList)
                {
                    
                    assetText = assetText + asset + "\n";
                    assets.Range.ListFormat.ApplyBulletDefault();
                }

                // Add it to the document 

                //assets.Format.
                assets.Range.ParagraphFormat.RightIndent = 35;
                assets.Range.ParagraphFormat.LeftIndent = 35;
                assets.Range.Text = assetText;
                assets.Range.ListFormat.ApplyBulletDefault();
                assets.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                assets.Range.ParagraphFormat.RightIndent = 35;
                assets.Range.ParagraphFormat.LeftIndent = 35;
                assets.Range.ListFormat.ApplyBulletDefault();
                assets.Range.InsertParagraph();
                //end of the bulleted list
              
                //this above stuff to assests need to be fixed
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                

                para10.Range.Text = returnNineteen();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Bold = 1;
                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle;
                para10.Range.Text = returnTwenty();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Bold = 1;
                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = returnTwentyOne();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Bold = 1;
                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle;
                para10.Range.Text = "Financial Results(Will Change based on the company)";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Bold = 0;
                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = returnTwentyTwo();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Bold = 1;
                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle;
                para10.Range.Text = returnTwentyThree();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Bold = 0;
                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = returnTwentyFour();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                //this entirley overwrites the previous footer rn
                Microsoft.Office.Interop.Word.Section wordSection2 = document.Content.Sections.Add(ref missing);
                //Get the footer range and add the footer details.  
                Microsoft.Office.Interop.Word.Range footerRange2 = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange2.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlack;
                footerRange2.Font.Size = 10;
                footerRange2.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                footerRange2.Text = "  The International Glossary of Business Valuation Terms defines “Going Concern” as “an ongoing operating business enterprise” and “Going Concern Value” as “the value of a business enterprise that is expected to continue to operate into the future.  The intangible elements of going-concern value result from factors such as having a trained workforce, an operational plant and the necessary licenses, systems and procedures in place.”";

                //starts actual page 4 not including header page
                para10.Range.Bold = 0;
                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle;
                para10.Range.Text = "Capitalization of Earnings Method";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                //dont need to reset bold unless you make it bold in the previous paragraph
                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = returnTwentyFive();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = returnTwentySix();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = returnTwentySeven();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = "The Build-Up Method utilizing the CRSP data is outlined below.";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                //addStuff(para3);
                //Create a 5X5 table and insert some dummy record  
                Table firstTable = document.Tables.Add(para10.Range, 1, 1, ref missing, ref missing);
                firstTable.Range.Text = "E(Ri)= RFR + ERP + RPS + RPU ";
                firstTable.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                firstTable.Range.ParagraphFormat.RightIndent = 35;
                firstTable.Range.ParagraphFormat.LeftIndent = 35;
                firstTable.Borders.Enable = 1;

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = "Where:";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = "E(Ri)	=	Expected return on investment";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 35;
                para10.Range.ParagraphFormat.LeftIndent = 35;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = "RFR = Risk - free rate of return";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 35;
                para10.Range.ParagraphFormat.LeftIndent = 35;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = "ERP = Equity risk premium";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 35;
                para10.Range.ParagraphFormat.LeftIndent = 35;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = "RPS = Risk premium for small stocks";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 35;
                para10.Range.ParagraphFormat.LeftIndent = 35;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = "RPU = Risk premium for unsystematic (company - specific) risk";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 35;
                para10.Range.ParagraphFormat.LeftIndent = 35;
                para10.Range.InsertParagraphAfter();

                para10.Range.Text = returnTwentyEight();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Text = returnTwentyNine();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();
                /*
                 * there is going to need to be a check here to see if the document runs onto a new page
                 * if so then the following paragraph follows
                 */

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle;
                para10.Range.Text = "Capitalization of Earnings Method – (Continued)";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = returnThirty();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = returnThirtyOne();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = "The Build-Up Method utilizing the RPR exhibits expressed as a formula is";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                /*
                 * this gets repeated but a little different
                 * 
                 */
                Table secondTable = document.Tables.Add(para10.Range, 1, 1, ref missing, ref missing);
                secondTable.Range.Text = "E(Ri) = RFR + ESRP + RPU ";
                secondTable.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                secondTable.Range.ParagraphFormat.RightIndent = 35;
                secondTable.Range.ParagraphFormat.LeftIndent = 35;
                secondTable.Borders.Enable = 1;

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = "Where:";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = "E(Ri)	=	Expected return on investment";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 35;
                para10.Range.ParagraphFormat.LeftIndent = 35;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = "RFR = Risk - free rate of return";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 35;
                para10.Range.ParagraphFormat.LeftIndent = 35;
                para10.Range.InsertParagraphAfter();

                

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = "ESRP = Equity / Size Risk Premium ";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 35;
                para10.Range.ParagraphFormat.LeftIndent = 35;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = "RPU = Risk premium for unsystematic (company - specific) risk";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 35;
                para10.Range.ParagraphFormat.LeftIndent = 35;
                para10.Range.InsertParagraphAfter();

                para10.Range.Text = returnThirtyTwo();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Text = returnThirtyThree();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Text = returnThirtyFour();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Text = returnThirtyFive();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Text = returnThirtySix();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();
                /*
                 * little bit of a different table but relativley the same
                 */

                /*
                 * there is going to need to be a check here to see if the document runs onto a new page
                 * if so then the following paragraph follows
                 */
                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle;
                para10.Range.Text = "Capitalization of Earnings Method – (Continued)";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = returnThirtySeven();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Text = returnThirtyEight();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle;
                para10.Range.Text = returnThirtyNine();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = returnForty();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle;
                para10.Range.Text = returnFortyOne();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = returnFortyTwo();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Text = returnFortyThree();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Text = returnFortyFour();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineSingle;
                para10.Range.Text = "Summary";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = returnFortyFive();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Bold = 1;
                para10.Range.Text = returnFortySix();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                //this will need to be a variable for the text
                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineDouble;
                para10.Range.Text = "$13,658,000";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Bold = 0;
                para10.Range.Underline = Microsoft.Office.Interop.Word.WdUnderline.wdUnderlineNone;
                para10.Range.Text = returnFortySeven();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

               
                para10.Range.Text = returnFortyEight();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                
                para10.Range.Text = returnFortyNine();
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.ParagraphFormat.RightIndent = 0;
                para10.Range.ParagraphFormat.LeftIndent = 0;
                para10.Range.InsertParagraphAfter();

                para10.Range.Text = "Sincerely,";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.InsertParagraphAfter();

                para10.Range.Bold = 1;
                para10.Range.Italic = 1;
                para10.Range.Text = "Apple Growth Partners";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.InsertParagraphAfter();

                //string imgPath = @"C:\Users\ddevine20\Downloads\CS470v2\CS470\WindowsFormsControlLibrary1\images\jason_sig.png";

                //Microsoft.Office.Interop.Word.InlineShape map = document.InlineShapes.AddPicture(imgPath);
                //map.Height = 350;
                //map.Width = 350;
                //map.Range.InsertPAfter();
                //map.Range.InsertAfter("Apple Growth Partners");

                para10.Range.InlineShapes.AddPicture(@"C:\Users\ddevine20\Downloads\CS470v2\CS470\WindowsFormsControlLibrary1\images\jason_sig.png");

                para10.Range.Bold = 0;
                para10.Range.Italic = 0;
                para10.Range.Text = "Jason R. Bogniard, MBA, ASA, CVA, EA";
                para10.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphJustify;
                para10.Range.InsertParagraphAfter();

                // row.Range.Text = "E(Ri)= RFR + ERP + RPS + RPU ";
                /*cell.Range.Font.Bold = 1;
                //other format properties goes here  
                cell.Range.Font.Name = "verdana";
                cell.Range.Font.Size = 10;
                //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                              
                cell.Shading.BackgroundPatternColor = WdColor.wdColorGray25;
                //Center alignment for the Header cells  */
                //row.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                // row.Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;




                ///bring up file save dialog to allow the user to save the file where ever
                SaveFileDialog SaveFileDialog1 = new SaveFileDialog();
                SaveFileDialog1.ShowDialog();
                String saveAs = SaveFileDialog1.FileName;
                object filename = saveAs;
                document.SaveAs2(ref filename);
                document.Close(ref missing, ref missing, ref missing);
                document = null;
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
                MessageBox.Show("Document created successfully !");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private string returnGeneric()
        {
            string str;
            str = "";
            return str;
        }

        private string returnFifty()
        {
            string str;
            str = "";
            return str;
        }
        private string returnFortyNine()
        {
            string str;
            str = "Furthermore, we have not performed an audit or appraisal of the underlying assets; therefore, matters might have come to our attention that would alter the amount reported above.  This report is not intended to determine the effects of such adjustments, but only to determine the calculated value of the Company based upon the information provided to us and our analysis and rationale as reported in this calculation report.  Additionally, we have no personal interest or bias in this calculation of value and our fee is not contingent upon our findings.";
            return str;
        }

        private string returnFortyEight()
        {
            string str;
            str = "We have no responsibility, but reserve the right, to update this report for events and circumstances occurring and/or the receipt of additional information received subsequent to the date of this report.";
            return str;
        }

        private string returnFortySeven()
        {
            string str;
            str = "This calculation report is subject to the Assumptions and Limiting Conditions found in Appendix B of this report and to the Valuation Analyst’s Representation/Certification found in Appendix C of this report.";
            return str;
        }

        private string returnFortySix()
        {
            string str;
            str = "Based on the review of data and information provided, we find a calculated value of a 100% Ownership Interest in ABC on a controlling, marketable basis as of December 31, 2018 to be:";
            return str;
        }

        private string returnFortyFive()
        {
            string str;
            str = "We have performed a calculation engagement, as that term is defined in the Statement on Standards for Valuation Services No. 1 (SSVS) of the American Institute of Certified Public Accountants, of a 100% controlling, marketable ownership interest in ABC Company as of December 31, 2018.  The resulting estimate of value has been calculated for corporate planning purposes for Mr. Smith and should not be used for any other purpose or by any other party for any purpose.  This calculation engagement was conducted in accordance with the SSVS.  The estimate of value that results from a calculation engagement is expressed as a calculation of value.";
            return str;
        }
        private string returnFortyFour()
        {
            string str;
            str = "Typically, owners of privately held companies have limited access to an active public market for the sale of their ownership interests. Without public market access, an owner’s ability to control the timing of potential gains, avoid losses, and minimize the opportunity cost associated with privately held investments is impaired. Several lack of marketability discount studies have been performed, which suggest that discounts for lack of marketability range from 9% to 45%. Per our discussions with Management, however, there are several ready buyers which would likely result in a relatively swift sale. As a result, we find, as of the valuation date, that a discount for lack of marketability is not appropriate.";
            return str;
        }

        private string returnFortyThree()
        {
            string str;
            str = "Although the above is cast in mathematical terms, it represents a synthesis of our judgment as a market assessment of the value of the Company.  Other factors bearing on risk and potential are reflected in our choice of multiples and yield.";
            return str;
        }
        private string returnFortyTwo()
        {
            string str;
            str = "An entity’s value is comprised of the market’s assessment of the factors of value.  The influence of each factor may vary among entities and for the same entity from year to year.  On a going concern basis, earnings power, whether expressed in an income or market approach, normally is given the predominant consideration of the major factors.  Based on the methods examined, we conclude that the fair market value of the Company, before considering discounts, is $13,658,000, on a fully marketable, controlling interest basis, giving greater weight to the value calculated under the income approach.";
            return str;
        }
        private string returnFortyOne()
        {
            string str;
            str = "Calculation Adjustments";
            return str;
        }
        private string returnForty()
        {
            string str;
            str = "We searched for transactions using the DealStats® private transaction (“DealStats”) database.  Using this database, we found a total of 45 transactions involving fitness centers, gyms and other similar establishments in the Deals Stats database (see Schedules I and J).  By comparison, the Company places above the upper quartile in terms of revenue, but near the median quartile in terms of profitability.  Due to the similar profitability compared to the guideline range, we select a multiple at the median of the guideline merged & acquired transaction range.  We applied this multiple to the five-year weighted average results for sales and gross profit.  The suggested equity value for the Company utilizing this method and database is $13,750,000 (rounded) per Schedule K.";
            return str;
        }
        private string returnThirtyNine()
        {
            string str;
            str = "Guideline Merged & Acquired Company Method";
            return str;
        }


        private string returnThirtyEight()
        {
            string str;
            str = "We utilize the mid-year discounting convention to reflect that cash flows are received more or less evenly throughout the year.  The mid-year discounting convention projects that cash flows will be received earlier, on average, then at the year-end.  We then adjust the suggested operating equity value by normalized net- working capital deficiency.  After adjustments, the capitalized income method suggests an equity value of $13,560,000, rounded (see Schedule H)."; 
            return str;
        }

        private string returnThirtySeven()
        {
            string str;
            str = "The capitalization of a single period earnings method is most appropriate when a company’s current level of earnings is at a sustainable level, with future growth expected to be relatively stable and at a modest rate. The stable earnings base used in our model is ABC’s five-year weighted-average adjusted net income.  The net income is adjusted for the normalization of expenses as described previously in this report.  We also allow for changes in working capital, depreciation and capital expenditures. Therefore, the benefit stream utilized represents net cash flow available to equity investors.  We find that the five-year weighted-average adjusted results are reflective of the sustainable earnings of ABC in the future.  The cash flow translates into a suggested equity value by applying an appropriate cost of capital.  The cost of capital is also known as a discount rate, discussed above."; 
            return str;
        }

        private string returnThirtySix()
        {
            string str;
            str = "Capitalization Rate – To convert the discount rate to a capitalization rate, we adjust for ABC’s expected long-term growth, which should reflect the anticipated long-term growth rates of the industry and the economy.  The capitalization rate must reflect the economic life of a business into perpetuity.  Therefore, estimating a long-term growth rate in excess of industry or economic levels may be unrealistic because it would assure ABC would perpetually gain market share.  Therefore, we apply a long-term growth rate of 2.5% to reflect real economic growth, inflationary growth, and industry growth.  The resulting mid-year capitalization multiplier is 6.416 times (see Schedule G).";
            return str;
        }
        private string returnThirtyFive()
        {
            string str;
            str = "The Build-Up Method, based on both the CRSP format data and data from the RPR exhibits, suggests a total cost of equity of 20.00%. As a result, we select a cost of equity of 20.0% as of December 31, 2018.";
            return str;
        }
        private string returnThirtyFour()
        {
            string str;
            str = "Based on the factors mentioned above, we select a risk premium for unsystematic (company-specific) risk of 500 basis points for ABC.  Accordingly, the Build-Up Method based on the RPR exhibits suggests a cost of equity for the Company of 20.00% (see Schedule G).";
            return str;
        }
        private string returnThirtyThree()
        {
            string str;
            str = "In determining an equity/size risk premium for ABC, we compared the Company to the portfolio most similar in size (the 25th portfolio) for each indicator based on MVIC, market value of equity, current year sales, five-year average, EBITDA, five-year net increase, and total assets. We utilize the average equity/size risk premium of the selected indicators, or 12.13%.";
            return str;
        }
        private string returnThirtyTwo()
        {
            string str;
            str = "The risk-free rate of return used in the RPR build-up method is also based on the yield on the 20-year Treasury bond on December 31, 2018, or 2.87%.";
            return str;
        }

        private string returnThirtyOne()
        {
            string str;
            str = "The RPR exhibits measure equity/size risk premiums for stocks listed on the New York Stock Exchange, NYSE Amex, and NASDAQ exchanges from 1963 through 2018 based on six benchmarks of company size. The benchmarks of company size include market value of equity, five-year average net income, MVIC, total assets, five-year average EBITDA, and sales.  For each benchmark, publicly traded companies are divided into 25 size-ranked portfolios with the 1st portfolio containing the largest companies as measured by the respective benchmark.  Certain companies such as financial-related companies and financially-distressed companies are excluded from the study.  The equity/size risk premiums are measured against a risk-free rate or the income return on 20-year U.S. Treasury bonds.";
            return str;
        }

        private string returnThirty()
        {
            string str;
            str = "An additional risk premium unsystematic (company-specific) risk is associated with an investment in ABC. Specific risks include key person risk, geographic concentration, high competition, industry risk, and size risk in excess of the small stock risk premium. It is our judgment that an additional risk premium of 500 basis points is appropriate at this time, suggesting a total cost of equity for the Company of 20.00% (see Schedule G).";
            return str;
        }

        private string returnTwentyNine()
        {
            string str;
            str = "The equity risk premium represents the extra return demanded by an average equity investor in excess of the risk-free rate. The CRSP data measures the returns of the S&P 500 against the income returns of 20-year U.S. Treasury bonds from 1926 through 2018.  The indicated equity risk premium is 6.91%.  The CRSP data also indicates that the stock of smaller companies has historically commanded higher returns than the stock of larger companies.  The risk premium for small stocks, measured by the smallest ten percent of stocks on the New York Stock Exchange, NYSE-Amex and NASDAQ exchanges (less than $322 million in market capitalization) is 5.22%.";
            return str;
        }

        private string returnTwentyEight()
        {
            string str;
            str = "The risk-free rate of return represents the return that an investor would require from a relatively riskless investment.  The risk-free rate of return used in our model is the yield on the 20-year Treasury bond on December 31, 2018, or 2.87%.  The risk-free rate was determined from data from the Federal Reserve Bank.";
            return str;
        }

        private string returnTwentySeven()
        {
            string str;
            str = "Our selected cost of equity is calculated by the Build-Up Method.  We have valued the Company on direct to equity cash flow basis.  Therefore, we utilize cost of equity data from the Duff & Phelps, LLC Cost of Capital Navigator 2019 which presents data in two formats, the CRSP Deciles Size Study (“CRSP”) and the Risk Premium Report Study (“RPR”).";
            return str;
        }
        private string returnTwentySix()
        {
            string str;
            str = "Discount rates vary depending on the industry, size of the subject company, and numerous other risk factors.  All else being equal, a higher discount rate lowers a value for a company.  Because we prepare our analysis on an equity basis, it is necessary to consider the Company’s cost of equity.  The Company’s cost of equity represents the rate of return that equity investors require for an investment in the Company.  We adopt a cost of equity of 20.0% for the Company based on the calculation explained below.";
            return str;
        }
        private string returnTwentyFive()
        {
            string str;
            str = "The capitalization of earnings method is a single-period income model, which values a business based upon a stable earnings base and capitalizes that benefit stream based on the appropriate risk adjusted rate of return.  The steps involved in using the capitalization of earnings method are: select a stable earnings base; select an appropriate capitalization rate; and capitalize the stable earnings base.  The capitalization of earnings method is presented on Schedule H and is discussed below.  The earnings base used in our model is the Company’s five-year weighted-average adjusted net income. We also allow for changes in working capital, depreciation and capital expenditures. Our net working capital calculation is based on a normalized level of 6.7% of sales as we feel the Company’s level of net working capital at year end is not reflective of the true net working capital requirement. This level is based on the RMA Annual Studies total officer compensation for Fitness and Recreational Sports Centers.";
            return str;
        }


        private string returnTwentyFour()
        {
            string str;
            str = "In order to determine the true earnings power of the Company, we found it necessary to adjust historically-reported earnings for certain items. We reviewed officer compensation and adjusted compensation based on the RMA Annual Studies total officer compensation for Fitness and Recreational Sports Centers. We add back non-recurring or non-operational expenses including amortization, and life insurance premiums. Finally, we normalized adjusted earnings for income taxes at a rate of 17.5% per Schedule D.";
            return str;
        }

        private string returnTwentyThree()
        {
            string str;
            str = "Earnings Adjustments (Will Change based on the company)";
            return str;
        }
        private string returnTwentyTwo()
        {
            string str;
            str = "The Company’s balance sheets are presented on Schedule A. The Company’s historical statements of income are presented in Schedule B and Schedule C.";
            return str;
        }
        private string returnTwentyOne()
        {
            string str;
            str = "ABC Company was incorporated on September 1, 2011 and is taxed as an S corporation. The Company is a Junior Olympic Volleyball Club with National, American and Regional level teams for age groups 10 to 18.The Company leases its facility, located at 123 Main Street, Beachwood, Ohio, from an unrelated party. Mr.Smith holds a 50 % ownership interest in the Company, with Mrs. Smith holding the remaining 50 %.";
            return str;
        }
        private string returnTwenty()
        {
            string str;
            str = "History and Nature of Company (Will Change based on the company)";


            return str;
        }

        private string returnNineteen()
        {
            string str;
            str = "The analysis included, but was not limited to, the above-mentioned factors." + Environment.NewLine + 

    "The report is based on historical and prospective financial information provided to us by management and other third - parties. While nothing during our analysis indicated that such information could not be relied upon, we take no responsibility for the underlying data utilized in this report. Had we audited the underlying data, matters may have come to our attention which would have resulted in our using amounts which differ from those provided.  Users of this calculation report should be aware that the calculations are based on future earnings potential, based upon facts known, as of the calculation date, which may or may not have subsequently been realized. Therefore, the actual results achieved during the projection period may vary from the projections used in this calculation, and the variations may be material." + Environment.NewLine +
    "The premise of value is going concern.  The liquidation premise of value was considered and rejected as not applicable, as the going - concern value results in a more appropriate value for the interest than the liquidation value, whether orderly or forced." + Environment.NewLine +


    "The income, market, and asset approaches to value were considered in this calculation. Specifically, we utilized the capitalization of earnings and guideline merged & acquired transaction methods. Our calculation of value reflects these findings, our judgment and knowledge of the marketplace, and our expertise. Apple Growth Partners’ staff, under the direct supervision of the lead appraiser on this engagement, assisted in performing research, populating models with data and providing other general assistance." + Environment.NewLine +


  "In performing our work, we were provided with and / or relied upon various sources of information, including items listed in Appendix A." + Environment.NewLine +
  "We have no present or contemplated financial interest in the Company for which value has been calculated. Our fees, for this valuation, are based upon our normal hourly billing rates and are, in no way, contingent upon the results of our findings. We have no responsibility, but reserve the right, to update this report for events and circumstances occurring subsequent to the date of this report.";
            return str;
        }
        private string returnEight()
        {
            string str;
            str = "“This type of report should be used to communicate the results of a calculation engagement (calculated value); it should not be used to communicate the results of a valuation engagement (conclusion of value) (Paragraph 73).”";
            return str;
        }

        private string returnNine()
        {
            string str;
            str = "A calculation report has certain requirements, as presented in Paragraphs 73 to 76 of SSVS.";
            return str;
        }

        private string returnTen()
        {
            string str;
            str = "The analysis and report are in conformance with the ethics standards and business valuation development guidelines of the AICPA, NACVA,^1  and ASA";
            return str;
        }

        private string returnEleven()
        {
            string str;
            str = "The analysis is also in conformance with various revenue rulings, including Revenue Ruling 59-60, which outline the approaches, methods, and factors to be considered in valuing closely held corporations for federal tax purposes.  Rev. Ruling 65-192 extended the concepts in Rev. Ruling 59-60 to income and other tax purposes as well as to business interests of any type." + Environment.NewLine +

            "The standard of value for this report is fair market value. Fair market value is defined in the International Glossary of Business Valuation Terms by the AICPA, the ASA, and the Canadian Institute of Charted Business Valuators, the NACVA, and the Institute of Business Appraisers as:";
            return str;
        }

        private string returnTwelve()
        {
            string str;
            str = "“The price, expressed in terms of cash equivalents, at which property would change hands between a hypothetical willing and able buyer and a hypothetical willing and able seller, acting at arm’s length in an open and unrestricted market, when neither is under compulsion to buy or sell and when both have reasonable knowledge of the relevant facts.”";
            return str;
        }

        private string returnThirteen()
        {
            string str;
            str = "Fair market value is also defined in Revenue Ruling 59-60 as:";
            return str;
        }

        private string returnFourteen()
        {
            string str;
            str = "“The price at which the property would change hands between a willing buyer and a willing seller when the former is not under any compulsion to buy and the latter is not under any compulsion to sell, both parties having reasonable knowledge of relevant facts.”";
            return str;
        }

        private string returnFifthteen()
        {
            string str;
            str = "Revenue Ruling 59-60 also defines the willing buyer and seller as hypothetical as follows:";
            return str;
        }

        private string returnSixthteen()
        {
            string str;
            str = "“Court decisions frequently state in addition that the hypothetical buyer and seller are assumed to be able, as well as willing, to trade and to be well informed about the property and concerning the market for such property.”";
            return str;
        }

        private string returnSeventeen()
        {
            string str;
            str = "Furthermore, fair market value assumes that the price is transacted in cash or cash equivalents.  Revenue Ruling 59-60, while used in tax valuations, is also used in many non-tax valuations.";
            return str;
        }

        private string returnEighteen()
        {
            string str;
            str = "In our calculation of value, we have considered the following relevant factors, which are specified in Revenue Ruling 59 - 60:";
            return str;
        }

        public string returnBullets()
        {
            string str;
            str = "•	The history and nature of the business" +
"•	The economic outlook of the United States and that of the specific industry in particular"+
"•	The book value of the subject company’s stock and the financial condition of the business"+
"•	The earning capacity of the company"+
"•	The dividend-paying capacity of the company"+
"•	Whether or not the company has goodwill or other intangible value"+
"•	Sales of the stock and size of the block of stock to be valued" +
"•	The market price of publicly traded stocks of corporations engaged in similar industries or lines of business";

            return str;
        }

    }
}
