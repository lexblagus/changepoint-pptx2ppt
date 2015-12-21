using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;

using System.IO;
using System.Net;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace pptx2ppt
{
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            lblDebug.Text = "";
            lblDebug.Text += "Initialized at " + DateTime.Now.ToLongTimeString() + "…" + "\n";

            // Local working folders
            //string inputFile = "C:\\Users\\Administrator\\Gallery\\Work\\Ornitorrinko\\PDTI-Sebrae\\pptx2ppt\\publish\\input-files\\convert-this.pptx";
            //string outputFile = "C:\\Users\\Administrator\\Gallery\\Work\\Ornitorrinko\\PDTI-Sebrae\\pptx2ppt\\publish\\output-files\\converted.ppt";
            string inputFile = "";
            string outputFile = "";

            // Set input parameters
            string inputSource = Request.QueryString["source"];

            if (inputSource != null)
            {
                // Default mode
                lblDebug.Text += "…in default mode" + "\n";
                lblDebug.Text += "From " + inputSource + "\n";

                // Download paths
                string workDir = HttpContext.Current.Server.MapPath(".");
                string inputDir = workDir + "\\input-files";
                string fileGuid = Guid.NewGuid().ToString();
                string inputPath = inputDir + "\\" + fileGuid + ".pptx";
                bool inputDirExists = Directory.Exists(inputDir);
                lblDebug.Text += "Download file: " + inputPath + "\n";
                
                if(inputDirExists){
                    // Download and save
                    lblDebug.Text += "Downloading and saving…" + "\n";
                    //Response.Flush();
                    bool isSaved = true;
                    try {
                        WebClient myWebClient = new WebClient();
                        myWebClient.DownloadFile(inputSource, inputPath);
                    }
                    catch (Exception err)
                    {
                        isSaved = false;
                        lblDebug.Text += "Error saving file. Technical details:" + "\n";
                        lblDebug.Text += err.ToString() + "\n";
                    }
                    if (isSaved) { 
                        lblDebug.Text += "File saved." + "\n";

                        // Convert
                        string outputFolder = "output-files";
                        string outputDir = workDir + "\\" + outputFolder;
                        bool outputDirExists = Directory.Exists(outputDir);
                        if (outputDirExists)
                        {
                            string outputPath = outputDir + "\\" + fileGuid + ".ppt";
                            lblDebug.Text += "Convert to: " + outputPath + "\n";

                            lblDebug.Text += "Converting…" + "\n";
                            //Response.Flush();
                            bool isConverted = true;

                            try
                            {
                                //Instantiate a Presentation object that represents a PPTX file
                                Presentation pres = new Presentation(inputPath);

                                //Saving the PPTX presentation to PPT format
                                pres.Save(outputPath, SaveFormat.Ppt);
                            }
                            catch (Exception err)
                            {
                                isConverted = false;
                                lblDebug.Text += "Error converting file. Technical details:" + "\n";
                                lblDebug.Text += err.ToString() + "\n";
                            }
                            if (true || isConverted)
                            {
                                lblDebug.Text += "File converted." + "\n";
                                
                                // Deliver file
                                string redirectTo = HttpContext.Current.Request.Url.AbsolutePath.Replace("Default", "") + outputFolder + "/" + fileGuid + ".ppt";
                                lblDebug.Text += "Deliver: " + redirectTo + "\n";
                                Response.Clear();
                                Response.Redirect(redirectTo);
                            }
                        }
                        else
                        {
                            lblDebug.Text += "Error: output folder does not exist." + "\n";
                        }
                    
                    }
                }else{
                    lblDebug.Text += "Error: download folder does not exist." + "\n";
                }
             }else{
                // Interactive mode (webform or querystring)
                lblDebug.Text += "…in interactive mode" + "\n";
                pnlInteractive.Visible = true;
                
                // Get folder definitions
                inputFile = Request.QueryString["input"] != null ? Request.QueryString["input"] : Request.Form["input"];
                outputFile = Request.QueryString["output"] != null ? Request.QueryString["output"] : Request.Form["output"];
                lblDebug.Text += "Input=" + inputFile + "\n";
                lblDebug.Text += "Output=" + outputFile + "\n";
                
                if (IsPostBack || (inputFile != null && outputFile != null))
                {
                    if (inputFile != null && outputFile != null && inputFile != "" && outputFile != "")
                    {
                        if (File.Exists(inputFile)) {
                            lblDebug.Text += "Converting…" + "\n";
                            //Response.Flush();
                            bool isConverted = true;

                            try {
                                //Instantiate a Presentation object that represents a PPTX file
                                Presentation pres = new Presentation(inputFile);

                                //Saving the PPTX presentation to PPT format
                                pres.Save(outputFile, SaveFormat.Ppt);
                            }
                            catch (Exception err)
                            {
                                isConverted = false;
                                lblDebug.Text += "Error converting file. Technical details:" + "\n";
                                lblDebug.Text += err.ToString() + "\n";
                            }
                            if (isConverted) { 
                                lblDebug.Text += "File converted." + "\n";
                            }
                            //Response.Flush();
                        }
                        else
                        {
                            lblDebug.Text += "Error: input file does not exist." + "\n";
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
}