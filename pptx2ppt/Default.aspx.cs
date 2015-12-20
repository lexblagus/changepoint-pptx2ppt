using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.IO;
using Aspose.Slides;
using Aspose.Slides.Export;



namespace pptx2ppt
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            lblDebug.Text = "";
            if (IsPostBack || (Request.QueryString["input"] != null && Request.QueryString["output"] != null))
            {
                lblDebug.Text += "Initializing at " + DateTime.Now.ToLongTimeString() + "…" + "\n";
                
                // Set input parameters
                string inputFile = Request.QueryString["input"];
                string outputFile = Request.QueryString["output"];
                lblDebug.Text += "Input=" + inputFile + "\n";
                lblDebug.Text += "Output=" + outputFile + "\n";
                
                if (inputFile != null && outputFile != null && inputFile != "" && outputFile != "")
                {
                    if (File.Exists(inputFile)) {

                        // Manual or debug
                        //string inputFile = "C:\\Users\\Administrator\\Gallery\\Work\\Ornitorrinko\\PDTI-Sebrae\\pptx2ppt\\publish\\input-files\\convert-this.pptx";
                        //string outputFile = "C:\\Users\\Administrator\\Gallery\\Work\\Ornitorrinko\\PDTI-Sebrae\\pptx2ppt\\publish\\output-files\\converted.ppt";

                        lblDebug.Text += "Converting…" + "\n";
                        Response.Flush();

                        //Instantiate a Presentation object that represents a PPTX file
                        Presentation pres = new Presentation(inputFile);

                        //Saving the PPTX presentation to PPT format
                        pres.Save(outputFile, SaveFormat.Ppt);

                        lblDebug.Text += "Converted." + "\n";
                        Response.Flush();
                    }
                    else
                    {
                        lblDebug.Text += "Error: input file does not exists." + "\n";
                    }
                }
                else
                {
                    lblDebug.Text += "No files selected." + "\n";
                }
                lblDebug.Text += "Finished!" + "\n";
            }else{
                lblDebug.Text += "Converter at your service." + "\n";
            }
        }
    }
}