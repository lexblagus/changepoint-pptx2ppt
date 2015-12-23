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
using System.Configuration;
using System.Data.SqlClient;
using System.Data;
using System.IO.Compression;

namespace pptx2ppt
{
    public partial class _Default : Page
    {
        private string confLogLevel     = "debug"; // none|info|debug
        private string confInputFolder  = "input-files"; // Relative to current folder. Use "\\" for folders, e.g.: files\\input
        private string confOutputFolder = "output-files"; // Relative to current folder. Use "\\" for folders, e.g.: files\\output
        private string confSQLhost      = "";
        private string confSQLdb        = "";
        private string confSQLuser      = "";
        private string confSQLpass = "";
        private bool confSimulation = false;

        protected void Page_Load(object sender, EventArgs e)
        {
            uiLog("clear","");
            uiLog("debug", "Page_Load()");
            uiLog("info", "Initialized at " + DateTime.Now.ToLongTimeString() + "…");

            ReadAppSettings();

            uiLog("debug", "confLogLevel=" + confLogLevel);
            uiLog("debug", "confInputFolder=" + confInputFolder);
            uiLog("debug", "confOutputFolder=" + confOutputFolder);
            uiLog("debug", "confSQLhost=" + confSQLhost);
            uiLog("debug", "confSQLdb=" + confSQLdb);
            uiLog("debug", "confSQLuser=" + confSQLuser);
            uiLog("debug", "confSQLpass=" + confSQLpass);

            // Set input parameters
            string inputSource = Request.QueryString["source"];
            uiLog("debug", "inputSource=" + inputSource);
            string inputFileId = Request.QueryString["FileId"];
            uiLog("debug", "inputFileId=" + inputFileId);
            string inputJobId = Request.QueryString["JobId"];
            uiLog("debug", "inputJobId=" + inputJobId);

            if (inputSource != null)
            {
                beginByAddress(inputSource);
            }else{
                if (inputFileId != null && inputJobId != null)
                {
                    beginIntelledox(inputFileId, inputJobId);
                }else{
                    confSimulation = true;
                    beginInteractive();
                }
                
            }
        }

        protected void beginInteractive()
        {
            uiLog("debug", "beginInteractive()");
            
            // Interactive mode (webform or querystring)
            uiLog("info", "…in interactive mode");
            pnlInteractive.Visible = true;

            // Get folder definitions
            string inputFile = Request.QueryString["input"] != null ? Request.QueryString["input"] : Request.Form["input"];
            string outputFile = Request.QueryString["output"] != null ? Request.QueryString["output"] : Request.Form["output"];
            uiLog("info", "Input=" + inputFile);
            uiLog("info", "Output=" + outputFile);

            if (IsPostBack || (inputFile != null && outputFile != null))
            {
                if (inputFile != null && outputFile != null && inputFile != "" && outputFile != "")
                {
                    if (File.Exists(inputFile))
                    {
                        convertPPTX2PPT(inputFile, outputFile);
                    }
                    else
                    {
                        uiLog("info", "Error: input file does not exist.");
                    }
                }
                else
                {
                    uiLog("info", "No files selected.");
                }
                uiLog("info", "Finished!");
            }
            else
            {
                uiLog("info", "Converter at your service.");
            }
        }

        protected void beginByAddress(string inputSource)
        {
            // Default mode
            uiLog("debug", "beginByAddress(\"" + inputSource + "\")");
            uiLog("info", "…in default mode");
            uiLog("info", "From " + inputSource);

            // Download paths
            string workDir = HttpContext.Current.Server.MapPath(".");
            string inputDir = workDir + "\\" + confInputFolder;
            string fileGuid = Guid.NewGuid().ToString();
            string inputPath = inputDir + "\\" + fileGuid + ".pptx";
            bool inputDirExists = Directory.Exists(inputDir);
            uiLog("info", "Download file: " + inputPath);
            uiLog("debug", "workDir=" + workDir);
            uiLog("debug", "inputDir=" + inputDir);
            uiLog("debug", "fileGuid=" + fileGuid);
            uiLog("debug", "inputPath=" + inputPath);

            if (inputDirExists)
            {
                // Download and save
                uiLog("info", "Downloading and saving…");
                if (confSimulation) { Response.Flush(); }
                bool isSaved = true;
                try
                {
                    WebClient myWebClient = new WebClient();
                    myWebClient.DownloadFile(inputSource, inputPath);
                }
                catch (Exception err)
                {
                    isSaved = false;
                    uiLog("info", "Error saving file. Technical details:");
                    uiLog("info", err.ToString());
                }
                if (isSaved)
                {
                    uiLog("info", "File saved.");
                    if (confSimulation) { Response.Flush(); }

                    // Convert
                    string outputFolder = confOutputFolder;
                    string outputDir = workDir + "\\" + outputFolder;
                    bool outputDirExists = Directory.Exists(outputDir);
                    if (outputDirExists)
                    {
                        string outputPath = outputDir + "\\" + fileGuid + ".ppt";
                        uiLog("info", "Convert to: " + outputPath);

                        uiLog("info", "Converting…");
                        bool isConverted = convertPPTX2PPT(inputPath, outputPath);
                        if (isConverted)
                        {
                            // Deliver file
                            string redirectTo = HttpContext.Current.Request.Url.AbsolutePath.Replace("Default", "") + outputFolder + "/" + fileGuid + ".ppt";
                            uiLog("info", "Deliver: " + redirectTo);
                            if (!confSimulation)
                            {
                                deliverFile(redirectTo, fileGuid + ".ppt");
                            }
                            else {
                                uiLog("info", "Running in simulation mode. Does not deliver the file");
                            }
                        }
                    }
                    else
                    {
                        uiLog("info", "Error: output folder does not exist.");
                    }

                }
            }
            else
            {
                uiLog("info", "Error: download folder does not exist.");
            }
        }

        protected void beginIntelledox(string FileId, string JobId)
        {
            uiLog("debug", "beginIntelledox(\"" + FileId + "\", \"" + JobId + "\") ");
            try {
                string strQuery = "select top 1 [DisplayName]+[Extension] as filename, [DocumentBinary] from [InfinitiDatabase].[dbo].[Document] where [DocumentId] = @FileId and [JobId] = @JobId";
                uiLog("debug", "strQuery=" + strQuery);
                SqlCommand cmd = new SqlCommand(strQuery);
                cmd.Parameters.Add("@FileId", SqlDbType.VarChar).Value = FileId;
                cmd.Parameters.Add("@JobId", SqlDbType.VarChar).Value = JobId;
                DataTable dt = GetData(cmd);
                if (dt != null)
                {
                    //download(dt);
                    string compressedPath = saveToDisk(dt);
                    if (compressedPath != "")
                    {
                        uiLog("debug", "Decompressing file");
                        string extractedPath = decompress(compressedPath);
                        uiLog("debug", "extractedPath=" + extractedPath);
                        if (extractedPath != "")
                        {
                            string friendlyName = Path.GetFileName(extractedPath);
                            string withoutExtension = Path.GetFileNameWithoutExtension(extractedPath);
                            string workDir = HttpContext.Current.Server.MapPath(".");
                            string outputDir = workDir + "\\" + confOutputFolder;
                            string outputPath = outputDir + "\\" + withoutExtension + ".ppt";
                            string redirToAddr = HttpContext.Current.Request.Url.AbsolutePath.Replace("Default", "") + confOutputFolder + "/" + withoutExtension + ".ppt";

                            uiLog("debug", "friendlyName=" + friendlyName);
                            uiLog("debug", "withoutExtension=" + withoutExtension);
                            uiLog("debug", "workDir=" + workDir);
                            uiLog("debug", "outputDir=" + outputDir);
                            uiLog("debug", "outputPath=" + outputPath);
                            uiLog("debug", "redirToAddr=" + redirToAddr);
                            
                            bool isConverted = convertPPTX2PPT(extractedPath, outputPath);
                            if (isConverted)
                            {
                                // Deliver file
                                uiLog("info", "Deliver: " + redirToAddr);
                                if (!confSimulation)
                                {
                                    deliverFile(redirToAddr, friendlyName);
                                }
                                else
                                {
                                    uiLog("info", "Running in simulation mode. Does not deliver the file");
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception err)
            {
                uiLog("info", "Error converting file. Technical details:");
                uiLog("info", err.ToString());
            }
        }

        private DataTable GetData(SqlCommand cmd)
        {
            uiLog("debug", "GetData(\"" + cmd.CommandText + "\") ");
            
            DataTable dt = new DataTable();
            String strConnString = System.Configuration.ConfigurationManager.ConnectionStrings["IntelledoxConnection"].ConnectionString;
            uiLog("debug", "strConnString=" + strConnString);
            SqlConnection con = new SqlConnection(strConnString);
            SqlDataAdapter sda = new SqlDataAdapter();
            cmd.CommandType = CommandType.Text;
            cmd.Connection = con;
            try
            {
                con.Open();
                sda.SelectCommand = cmd;
                sda.Fill(dt);
                return dt;
            }
            catch (Exception err)
            {
                uiLog("info", "Error getting file from database. Technical details:");
                uiLog("info", err.ToString());
                return null;
            }
            finally
            {
                con.Close();
                sda.Dispose();
                con.Dispose();
            }
        }

        private string saveToDisk(DataTable dt)
        {
            uiLog("debug", "saveToDisk() ");

            // Download paths
            string workDir = HttpContext.Current.Server.MapPath(".");
            string inputDir = workDir + "\\" + confInputFolder;
            bool inputDirExists = Directory.Exists(inputDir);
            string inputPath = inputDir + "\\" + dt.Rows[0]["filename"].ToString() + ".gz";
            
            uiLog("info", "Download file: " + inputPath);
            uiLog("debug", "workDir=" + workDir);
            uiLog("debug", "inputDir=" + inputDir);
            uiLog("debug", "inputDirExists=" + inputDirExists);
            uiLog("debug", "inputPath=" + inputPath);

            if (inputDirExists)
            {
                Byte[] fileBytes = (Byte[])dt.Rows[0]["DocumentBinary"];
                File.WriteAllBytes(inputPath, fileBytes);
                uiLog("info", "File from DB has been saved into the filesystem");
                return inputPath;
            }
            else
            {
                uiLog("info", "Error: input folder does not exist.");
                return "";
            }
            /*
            Response.Buffer = true;
            Response.Charset = "";
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            //Response.ContentType = dt.Rows[0]["ContentType"].ToString();
            Response.AddHeader("content-disposition", "attachment;filename=" + dt.Rows[0]["filename"].ToString());
            Response.BinaryWrite(bytes);
            Response.Flush();
            Response.End();
            */
        }

        private string decompress(string inputFile)
        {
            uiLog("debug", "decompress(\"" + inputFile + "\")");

            FileInfo fileToDecompress = new FileInfo(inputFile);

            using (FileStream originalFileStream = fileToDecompress.OpenRead())
            {
                string currentFileName = fileToDecompress.FullName;
                string newFileName = currentFileName.Remove(currentFileName.Length - fileToDecompress.Extension.Length);

                using (FileStream decompressedFileStream = File.Create(newFileName))
                {
                    using (GZipStream decompressionStream = new GZipStream(originalFileStream, CompressionMode.Decompress))
                    {
                        decompressionStream.CopyTo(decompressedFileStream);
                        uiLog("info", "Decompressed: " + newFileName);
                        return newFileName;
                    }
                }
            }
            return "";
        }

        private void download(DataTable dt)
        {
            uiLog("debug", "download(bytes) ");

            Byte[] bytes = (Byte[])dt.Rows[0]["DocumentBinary"];
            Response.Buffer = true;
            Response.Charset = "";
            Response.Cache.SetCacheability(HttpCacheability.NoCache);
            //Response.ContentType = dt.Rows[0]["ContentType"].ToString();
            Response.AddHeader("content-disposition", "attachment;filename=" + dt.Rows[0]["filename"].ToString());
            Response.BinaryWrite(bytes);
            Response.Flush();
            Response.End();
        }

        protected bool convertPPTX2PPT(string inputFile, string outputFile)
        {
            uiLog("debug", "convertPPTX2PPT(\"" + inputFile + "\", \"" + outputFile + "\")");
            uiLog("info", "Converting…");
            if (confSimulation) { Response.Flush(); }
            bool isConverted = true;

            try
            {
                //Instantiate a Presentation object that represents a PPTX file
                Presentation pres = new Presentation(inputFile);

                //Saving the PPTX presentation to PPT format
                pres.Save(outputFile, SaveFormat.Ppt);
            }
            catch (Exception err)
            {
                isConverted = false;
                uiLog("info", "Error converting file. Technical details:");
                uiLog("info", err.ToString());

            }
            if (isConverted)
            {
                uiLog("info", "File converted.");
            }
            if (confSimulation) { Response.Flush(); }
            return isConverted;
            
        }

        protected void deliverFile(string redirToAddr , string friendlyName)
        {
            uiLog("debug", "deliverFile(\"" + redirToAddr + "\", \"" + friendlyName + "\")");
            Response.Clear();
            Response.AppendHeader("Content-Disposition", "attachment;filename=" + friendlyName);
            Response.Redirect(redirToAddr);
        }

        protected void uiLog(string mode, string message)
        {
            if (
                (confLogLevel=="info" && (mode=="info")) ||
                (confLogLevel=="debug" && (mode=="debug" || mode=="info"))
            ){
                lblDebug.Text += mode.ToUpper() + ' ' + message + "\n";
            }
            else if (mode=="clear")
            {
                lblDebug.Text = "";
            }
            
        }

        protected void ReadAppSettings()
        {
            uiLog("debug", "ReadAppSettings()");
            uiLog("info", "Reading application settings from web.config");
            try
            {
                System.Configuration.Configuration rootWebConfig1 = System.Web.Configuration.WebConfigurationManager.OpenWebConfiguration(null);
                uiLog("debug", "rootWebConfig1.FilePath=" + rootWebConfig1.FilePath);
                uiLog("debug", "rootWebConfig1.AppSettings.Settings.Count=" + rootWebConfig1.AppSettings.Settings.Count);
                if (rootWebConfig1.AppSettings.Settings.Count > 0)
                {
                    uiLog("debug", "Loading each setting");
                    confLogLevel = rootWebConfig1.AppSettings.Settings["ppt2ppt_logLevel"].Value.ToString();
                    confInputFolder = rootWebConfig1.AppSettings.Settings["ppt2ppt_inputFolder"].Value.ToString();
                    confOutputFolder = rootWebConfig1.AppSettings.Settings["ppt2ppt_outputFolder"].Value.ToString();
                    confSQLhost = rootWebConfig1.AppSettings.Settings["ppt2ppt_SQLhost"].Value.ToString();
                    confSQLdb = rootWebConfig1.AppSettings.Settings["ppt2ppt_SQLdb"].Value.ToString();
                    confSQLuser = rootWebConfig1.AppSettings.Settings["ppt2ppt_SQLuser"].Value.ToString();
                    confSQLpass = rootWebConfig1.AppSettings.Settings["ppt2ppt_SQLpass"].Value.ToString();
                    confSimulation = rootWebConfig1.AppSettings.Settings["ppt2ppt_Simulation"].Value.ToString() == "true";
                }
                else
                {
                    uiLog("info", "Error: no key setting at " + rootWebConfig1.FilePath);
                }
            }
            catch (Exception e)
            {
                uiLog("info", "Error reading application settings: "+e.Message);
            }
            uiLog("debug", "Finished reading settings");
        }

    }
}