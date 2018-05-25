using ExcelDataReader;
using Google;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v2;
using Google.Apis.Logging;
using Google.Apis.Services;
using Google.Apis.Upload;
using System;
using System.Configuration;
using System.Data;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ConsoleAPI
{

    /// <summary>
    /// A sample for the Drive API. This samples demonstrates resumable excel upload    
    /// </summary>
    class Program
    {
        static Program()
        {
            Logger = ApplicationContext.Logger.ForType<ResumableUpload<Program>>();
        }

        //Full path of the file you want to upload
        //private const string UploadFileName = @"FILE_TO_UPLOAD\Horse_Details_Sequence_update.xlsx";
        private string UploadFileName = ConfigurationManager.AppSettings["FilePath"];

        //File type
        private const string ContentType = @"application/vnd.ms-excel";

        /// <summary>The logger instance.</summary>
        private static readonly ILogger Logger;

        /// <summary>The Drive API scopes.</summary>
        private static readonly string[] Scopes = new[] { DriveService.Scope.DriveFile, DriveService.Scope.Drive };

        /// <summary>The file which was uploaded</summary>
        private static Google.Apis.Drive.v2.Data.File uploadedFile;

        static void Main(string[] args)
        {
            Console.WriteLine("Google Drive API Sample");

            try
            {
                new Program().Run().Wait();
            }
            catch (AggregateException ex)
            {
                foreach (var e in ex.InnerExceptions)
                {
                    Console.WriteLine("ERROR: " + e.Message);
                }
            }

            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
        }

        private async Task Run()
        {
            GoogleWebAuthorizationBroker.Folder = "Drive.Sample";
            UserCredential credential;
            using (var stream = new System.IO.FileStream("client_secrets.json",
                System.IO.FileMode.Open, System.IO.FileAccess.Read))
            {
                credential = await GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets, Scopes, "user", CancellationToken.None);
            }

            // Create the service.
            var service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = "Drive API Sample",
            });

            await UploadFileAsync(service);

            // Upload succeeded
            Console.WriteLine("\"{0}\" was uploaded successfully", uploadedFile.Title);
        }

        /// <summary>Uploads file asynchronously.</summary>
        private Task<IUploadProgress> UploadFileAsync(DriveService service)
        {
            var title = UploadFileName;
            if (title.LastIndexOf('\\') != -1)
            {
                title = title.Substring(title.LastIndexOf('\\') + 1);
            }

            //Application xlApp = new Application();

            ////if (xlApp == null)
            ////{
            ////    //MessageBox.Show("Excel is not properly installed!!");
            ////    return;
            ////}


            ////xlApp.DisplayAlerts = false;
            //string filePath = @"D:\MVC_TestProject\API\ConsoleAPI\ConsoleAPI\bin\Debug\FILE_TO_UPLOAD\Test.xlsx";
            //Workbook xlWorkBook = xlApp.Workbooks.Open(filePath, 0, false, 5, "", "", false, XlPlatform.xlWindows, "", true, false, 0, true, false, false);
            ////Microsoft.Office.Interop.Excel.Sheets worksheets = xlWorkBook.Worksheets;

            //Worksheet wkSheet = new Worksheet();
            //xlApp.DisplayAlerts = false;
            ////for (int i = xlApp.ActiveWorkbook.Worksheets.Count; i > 0; i--)
            //for (int i = xlApp.ActiveWorkbook.Worksheets.Count; i > 0; i--)
            //{
            //    //wkSheet = (Worksheet)xlApp.ActiveWorkbook.Worksheets[i];
            //    wkSheet = xlWorkBook.Worksheets[i];
            //    if (wkSheet.Name != "Table 2")
            //    {
            //        wkSheet.Delete();
            //    }
            //}
            //xlApp.DisplayAlerts = true;

            //xlWorkBook.Save();
            //xlWorkBook.Close();

            //releaseObject(wkSheet);
            //releaseObject(xlWorkBook);
            //releaseObject(xlApp);

            var uploadStream = new System.IO.FileStream(UploadFileName, System.IO.FileMode.Open,
                    System.IO.FileAccess.Read);

            var insert = service.Files.Insert(new Google.Apis.Drive.v2.Data.File { Title = title }, uploadStream, ContentType);

            insert.ChunkSize = FilesResource.InsertMediaUpload.MinimumChunkSize * 2;
            insert.ProgressChanged += Upload_ProgressChanged;
            insert.ResponseReceived += Upload_ResponseReceived;

            var task = insert.UploadAsync();

            task.ContinueWith(t =>
            {
                //this code will be called if the upload fails
                Console.WriteLine("Upload Failed. " + t.Exception);
            }, TaskContinuationOptions.NotOnRanToCompletion);
            task.ContinueWith(t =>
            {
                Logger.Debug("Closing the stream");
                uploadStream.Dispose();
                Logger.Debug("The stream was closed");
            });

            return task;
        }

        #region Progress and Response changes

        static void Upload_ProgressChanged(IUploadProgress progress)
        {
            Console.WriteLine(progress.Status + " " + progress.BytesSent);
        }

        static void Upload_ResponseReceived(Google.Apis.Drive.v2.Data.File file)
        {
            uploadedFile = file;
        }

        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                //MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        #endregion
    }
}
