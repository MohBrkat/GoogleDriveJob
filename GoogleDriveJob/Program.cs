using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Util.Store;
using GoogleDriveJob.DAL;
using GoogleDriveJob.Helpers;
using GoogleDriveJob.Models;
using Log4NetLibrary;
using Microsoft.Extensions.Configuration;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using File = Google.Apis.Drive.v3.Data.File;

namespace GoogleDriveJob
{
    public static class Program
    {
        // If modifying these scopes, delete your previously saved credentials
        // at ~/.credentials/drive-dotnet-quickstart.json
        static string[] Scopes = { DriveService.Scope.Drive,
                               DriveService.Scope.DriveFile };
        static string ApplicationName = "Jifiti Google Cloud Reporting";
        static DriveService service;
        private static IConfiguration _iconfiguration;

        static async System.Threading.Tasks.Task Main(string[] args)
        {
            try
            {
                GetAppSettingsFile();

                InintializeGoogleDriveService();

                var clientIdsString = _iconfiguration.GetSection("ClientIds").Value;
                List<string> clientIdList = clientIdsString.Split(',').ToList();
                foreach (var clientId in clientIdList)
                {
                    await GenerateTotalRevenueReportAsync(clientId);
                    await GenerateNewUsersReportAsync(clientId);
                    await GenerateListsCreatedReportAsync(clientId);
                    await GenerateListsProductsReportAsync(clientId);
                    await GeneratePopularProductsReportAsync(clientId);
                }

                SaveCurrentDateToFile();
            }
            catch (Exception e)
            {
                //string emails = _iconfiguration?.GetSection("JifitiEmails")?.Value;
                //new RegistryDAL(_iconfiguration).CreateEmailLog(emails, "Failed GDS Job error: " + e.Message, "Failed during running GDS");
                new SendEmailHelper(_iconfiguration).SendEmail(e, "Failed during running GDS");
                Logger.Error(e.Message);
            }
        }

        #region Popular Products Report
        private static async Task GeneratePopularProductsReportAsync(string clientId)
        {
            List<PopularProductsModel> popularProductsData = new List<PopularProductsModel>();
            List<IList<Object>> objNewRecords = new List<IList<Object>>();

            string reportName = _iconfiguration.GetSection("PopularProductsReportName").Value;
            string extension = "csv";
            string contentType = "text/csv";
            string fileName = $"{clientId}{reportName}.{extension}";
            string filePath = GetFilePath(fileName);
            DateTime from = GetLastUsedDateFile();
            DateTime to;

            bool includeHeaders = true;
            if (!System.IO.File.Exists(filePath))
            {
                Logger.Info("File Doesn't Exist, Get all Data from Date : 2018-3-1");
                var reportInitialDate = _iconfiguration.GetSection("ReportsInitialDate").Value;

                from = Convert.ToDateTime(reportInitialDate);
                to = DateTime.Now.AbsoluteEnd();

                popularProductsData = GetPopularProductsDataByClientId(clientId, from, to);
            }
            else
            {
                includeHeaders = false;
                //from = DateTime.Now.AbsoluteStart();
                to = DateTime.Now.AbsoluteEnd();

                popularProductsData = GetPopularProductsDataByClientId(clientId, from, to);

                foreach (var t in popularProductsData)
                {
                    IList<Object> obj = new List<Object>
                    {
                        t.SKU,
                        t.ProductTitle,
                        t.Category,
                        t.Price,
                        t.ItemsAdded,
                        t.ItemsBought
                    };
                    objNewRecords.Add(obj);
                }
            }

            await ExportAndUploadToDriveAsync(popularProductsData, objNewRecords, extension, contentType, fileName, includeHeaders);
        }

        static List<PopularProductsModel> GetPopularProductsDataByClientId(string clientId, DateTime from, DateTime to)
        {
            Logger.Info("Get Popular Products Data for Client Id " + clientId);
            var registryDAL = new RegistryDAL(_iconfiguration);
            int client = Convert.ToInt32(clientId);
            var popularProductsModels = registryDAL.GetPopularProductsData(client, from, to);
            return popularProductsModels;
        }

        #endregion

        #region List Products Report
        private static async Task GenerateListsProductsReportAsync(string clientId)
        {
            List<ListProductsModel> listProductsData = new List<ListProductsModel>();
            List<IList<Object>> objNewRecords = new List<IList<Object>>();

            string reportName = _iconfiguration.GetSection("ListProductsReportName").Value;
            string extension = "csv";
            string contentType = "text/csv";
            string fileName = $"{clientId}{reportName}.{extension}";
            string filePath = GetFilePath(fileName);
            DateTime from = GetLastUsedDateFile();
            DateTime to;

            bool includeHeaders = true;
            if (!System.IO.File.Exists(filePath))
            {
                Logger.Info("File Doesn't Exist, Get all Data from Date : 2018-3-1");
                var reportInitialDate = _iconfiguration.GetSection("ReportsInitialDate").Value;

                from = Convert.ToDateTime(reportInitialDate);
                to = DateTime.Now.AbsoluteEnd();

                listProductsData = GetListProductsDataByClientId(clientId, from, to);
            }
            else
            {
                includeHeaders = false;
                //from = DateTime.Now.AbsoluteStart();
                to = DateTime.Now.AbsoluteEnd();

                listProductsData = GetListProductsDataByClientId(clientId, from, to);

                foreach (var t in listProductsData)
                {
                    IList<Object> obj = new List<Object>
                    {
                        t.InsertDate,
                        t.ListProductID,
                        t.ListID,
                        t.ProductID,
                        t.ListProductStatus,
                        t.Quantity,
                        t.Price
                    };
                    objNewRecords.Add(obj);
                }
            }

            await ExportAndUploadToDriveAsync(listProductsData, objNewRecords, extension, contentType, fileName, includeHeaders);
        }

        static List<ListProductsModel> GetListProductsDataByClientId(string clientId, DateTime from, DateTime to)
        {
            Logger.Info("Get List Products Data for Client Id " + clientId);
            var registryDAL = new RegistryDAL(_iconfiguration);
            int client = Convert.ToInt32(clientId);
            var listProductsModels = registryDAL.GetListProductsData(client, from, to);
            return listProductsModels;
        }

        #endregion

        #region Total Revenue Report
        private static async System.Threading.Tasks.Task GenerateTotalRevenueReportAsync(string clientId)
        {
            List<TotalRevenueModel> totalRevenueData = new List<TotalRevenueModel>();
            List<IList<Object>> objNewRecords = new List<IList<Object>>();

            string reportName = _iconfiguration.GetSection("TotalRevenueReportName").Value;
            string extension = "csv";
            string contentType = "text/csv";
            string fileName = $"{clientId}{reportName}.{extension}";
            string filePath = GetFilePath(fileName);
            DateTime from = GetLastUsedDateFile();
            DateTime to;

            bool includeHeaders = true;
            if (!System.IO.File.Exists(filePath))
            {
                Logger.Info("File Doesn't Exist, Get all Data from Date : 2018-3-1");
                var reportInitialDate = _iconfiguration.GetSection("ReportsInitialDate").Value;

                from = Convert.ToDateTime(reportInitialDate);
                to = DateTime.Now.AbsoluteEnd();

                totalRevenueData = GetTotalRevenueDataByClientId(clientId, from, to);
            }
            else
            {
                includeHeaders = false;
                //from = DateTime.Now.AbsoluteStart();
                to = DateTime.Now.AbsoluteEnd();

                totalRevenueData = GetTotalRevenueDataByClientId(clientId, from, to);

                foreach (var t in totalRevenueData)
                {
                    IList<Object> obj = new List<Object>
                    {
                        t.InsertDate,
                        t.OrderId,
                        t.BuyerID,
                        t.OrderTransaction,
                        t.SourceType,
                        t.DeliveryOption,
                        t.SKU,
                        t.ProductPrice,
                        t.TotalRevenue
                    };
                    objNewRecords.Add(obj);
                }
            }

            await ExportAndUploadToDriveAsync(totalRevenueData, objNewRecords, extension, contentType, fileName, includeHeaders);
        }

        static List<TotalRevenueModel> GetTotalRevenueDataByClientId(string clientId, DateTime from, DateTime to)
        {
            Logger.Info("Get Total Revenue Data for Client Id " + clientId);
            var registryDAL = new RegistryDAL(_iconfiguration);
            int client = Convert.ToInt32(clientId);
            var totalRevenueModels = registryDAL.GetTotalRevenueData(client, from, to);
            return totalRevenueModels;
        }
        #endregion

        #region New Users Report
        private static async System.Threading.Tasks.Task GenerateNewUsersReportAsync(string clientId)
        {
            List<NewUsersModel> newUsersData = new List<NewUsersModel>();
            List<IList<Object>> objNewRecords = new List<IList<Object>>();

            string reportName = _iconfiguration.GetSection("NewUsersReportName").Value;
            var extension = "csv";
            var contentType = "text/csv";
            var fileName = $"{clientId}{reportName}.{extension}";
            string filePath = GetFilePath(fileName);
            DateTime from = GetLastUsedDateFile();
            DateTime to;

            bool includeHeaders = true;
            if (!System.IO.File.Exists(filePath))
            {
                Logger.Info("File Doesn't Exist, Get all Data from Date : 2018-3-1");
                var reportInitialDate = _iconfiguration.GetSection("ReportsInitialDate").Value;

                from = Convert.ToDateTime(reportInitialDate);
                to = DateTime.Now.AbsoluteEnd();

                newUsersData = GetNewUsersDataByClientId(clientId, from, to);
            }
            else
            {
                includeHeaders = false;
                //from = DateTime.Now.AbsoluteStart();
                to = DateTime.Now.AbsoluteEnd();

                newUsersData = GetNewUsersDataByClientId(clientId, from, to);

                foreach (var user in newUsersData)
                {
                    IList<Object> obj = new List<Object>
                    {
                        user.InsertDate,
                        user.UserId,
                        user.SourceType
                    };

                    objNewRecords.Add(obj);
                }
            }

            await ExportAndUploadToDriveAsync(newUsersData, objNewRecords, extension, contentType, fileName, includeHeaders);
        }

        private static List<NewUsersModel> GetNewUsersDataByClientId(string clientId, DateTime from, DateTime to)
        {
            Logger.Info("Get New Users Data for Client Id " + clientId);
            var registryDAL = new RegistryDAL(_iconfiguration);
            int client = Convert.ToInt32(clientId);
            var newUsersModels = registryDAL.GetNewUsersData(client, from, to);
            return newUsersModels;
        }

        #endregion

        #region Lists Created Report
        private static async System.Threading.Tasks.Task GenerateListsCreatedReportAsync(string clientId)
        {
            List<ListsCreatedModel> listsCreatedData = new List<ListsCreatedModel>();
            List<IList<Object>> objNewRecords = new List<IList<Object>>();

            string reportName = _iconfiguration.GetSection("ListsCreatedReportName").Value;
            var extension = "csv";
            var contentType = "text/csv";
            var fileName = $"{clientId}{reportName}.{extension}";
            string filePath = GetFilePath(fileName);
            DateTime from = GetLastUsedDateFile();
            DateTime to;

            bool includeHeaders = true;
            if (!System.IO.File.Exists(filePath))
            {
                var reportInitialDate = _iconfiguration.GetSection("ReportsInitialDate").Value;
                Logger.Info($"File Doesn't Exist, Get all Data from Date : {reportInitialDate}");

                from = Convert.ToDateTime(reportInitialDate);
                to = DateTime.Now.AbsoluteEnd();

                listsCreatedData = GetListsCreatedDataByClientId(clientId, from, to);
            }
            else
            {
                includeHeaders = false;
                //from = DateTime.Now.AbsoluteStart();
                to = DateTime.Now.AbsoluteEnd();

                listsCreatedData = GetListsCreatedDataByClientId(clientId, from, to);

                foreach (var list in listsCreatedData)
                {
                    IList<Object> obj = new List<Object>
                    {
                        list.CreatedDate,
                        list.EventDate,
                        list.LastPurchasedDate,
                        list.ListId,
                        list.UniqueURL,
                        list.Products
                    };

                    objNewRecords.Add(obj);
                }
            }

            await ExportAndUploadToDriveAsync(listsCreatedData, objNewRecords, extension, contentType, fileName, includeHeaders);
        }

        private static List<ListsCreatedModel> GetListsCreatedDataByClientId(string clientId, DateTime from, DateTime to)
        {
            Logger.Info("Get Lists Created Data for Client Id " + clientId);
            var registryDAL = new RegistryDAL(_iconfiguration);
            int client = Convert.ToInt32(clientId);
            var listsCreatedModels = registryDAL.GetListsCreatedData(client, from, to);
            return listsCreatedModels;
        }
        #endregion


        #region Google Drive Services
        private static void InintializeGoogleDriveService()
        {
            Logger.Info("Inintialize Google Drive Service");

            UserCredential credential;

            using (var stream =
                new FileStream("credentials.json", FileMode.Open, FileAccess.Read))
            {
                // The file token.json stores the user's access and refresh tokens, and is created
                // automatically when the authorization flow completes for the first time.
                string credPath = "token.json";
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    Scopes,
                    $"{Environment.UserName}.drive",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)).Result;
                Console.WriteLine("Credential file saved to: " + credPath);
            }

            // Create Drive API service.
            service = new DriveService(new BaseClientService.Initializer()
            {
                HttpClientInitializer = credential,
                ApplicationName = ApplicationName,
            });
        }

        private static async System.Threading.Tasks.Task ExportAndUploadToDriveAsync<T>(List<T> dataList, List<IList<Object>> writeValues, string extension, string contentType, string fileName, bool includeHeaders = true)
        {
            //Generate and Save Excel File locally
            Logger.Info("Generating Excel File");
            List<List<T>> splittedData = ExportToExcelHelper.Split(dataList, 1000000);
            List<byte> data = new List<byte>();
            foreach (var dt in splittedData)
            {
                var r = ExportToExcelHelper.ExportToExcel(dt, extension, includeHeaders).ToList();
                data.AddRange(r);
            }

            Logger.Info("Saving Excel File in ReportFiles folder");
            string filePath = GetFilePath(fileName);

            using (var stream = new FileStream(filePath, FileMode.Append))
            {
                stream.Write(data.ToArray(), 0, data.ToArray().Length);
            }

            //System.IO.File.WriteAllBytes(filePath, data.ToArray());

            //Google Drive Folder Name
            string folderName = _iconfiguration.GetSection("GoogleDriveFolderName").Value;
            IList<File> folders = GetFolders();

            var folder = folders.FirstOrDefault(a => a.Name == folderName);

            if (folder == null)
            {
                folder = service.CreateFolder(folderName);
            }

            if (folder != null)
            {
                Logger.Info($"Folder: {folder.Name} Found");
                var fs = GetFilesByName(Path.GetFileNameWithoutExtension(fileName));
                if (fs.Count == 1)
                {

                    //UpdateFile(filePath, folder, fs.FirstOrDefault(), contentType);
                    await UpdateFileUsingGoogleSheetAsync<T>(fs.FirstOrDefault(), writeValues);
                }
                else
                {
                    CreateFile(filePath, folder, contentType);
                }
            }
        }

        private static async System.Threading.Tasks.Task UpdateFileUsingGoogleSheetAsync<T>(File file, List<IList<Object>> dataList)
        {
            var sheetId = file.Id;
            await GoogleSheet.WriteAsync(dataList, sheetId);
        }

        private static void CreateFile(string filePath, File folder, string contentType)
        {
            string fileNameWithExtesntion = Path.GetFileName(filePath);

            string fileName = Path.GetFileNameWithoutExtension(fileNameWithExtesntion);

            var fileMetadata = new File()
            {
                Name = fileName,
                Parents = new List<string> { folder.Id },
                MimeType = "application/vnd.google-apps.spreadsheet"
            };

            var mimeType = GetMimeType(fileNameWithExtesntion);

            Logger.Info($"Creating File: {fileMetadata.Name} in Folder: {folder.Name}");
            FilesResource.CreateMediaUpload request;
            using (var stream = new System.IO.FileStream(filePath,
                                    System.IO.FileMode.Open))
            {
                request = service.Files.Create(
                    fileMetadata, stream, contentType);
                request.Fields = "id,name";
                request.KeepRevisionForever = true;
                request.Upload();
            }
            var fileCreated = request.ResponseBody;
            if (!string.IsNullOrEmpty(fileCreated.Id))
            {
                Logger.Info($"File {fileCreated.Name} Created");
            }
            Console.WriteLine($"File {fileCreated.Name} Created");
        }

        private static IList<File> GetFolders()
        {
            Logger.Info("Getting Google Drive Folders");

            // Define parameters of request.
            FilesResource.ListRequest listRequest = service.Files.List();
            listRequest.PageSize = 10;
            listRequest.Q = "mimeType = 'application/vnd.google-apps.folder' and trashed = false";
            listRequest.Fields = "nextPageToken, files(id, name, parents)";
            listRequest.Spaces = "drive";
            // List files.
            return listRequest.Execute().Files;
        }

        private static IList<File> GetFilesByName(string fileName)
        {
            Logger.Info("Getting Google Drive Files");

            // Define parameters of request.
            FilesResource.ListRequest listRequest = service.Files.List();
            listRequest.PageSize = 10;
            listRequest.Q = "mimeType != 'application/vnd.google-apps.folder' and trashed = false and name='" + fileName + "'";
            listRequest.Fields = "nextPageToken, files(id, name, parents)";
            listRequest.Spaces = "drive";
            // List files.
            return listRequest.Execute().Files;
        }

        private static File CreateFolder(this DriveService service, string folderName)
        {
            Logger.Info("Creating Folder: " + folderName);
            var fileMetadata = new Google.Apis.Drive.v3.Data.File()
            {
                Name = folderName,
                MimeType = "application/vnd.google-apps.folder"
            };
            var request = service.Files.Create(fileMetadata);
            request.Fields = "id,name";
            var file = request.Execute();
            if (!string.IsNullOrEmpty(file.Id))
            {
                Logger.Info($"Folder {file.Name} Created");
            }

            Console.WriteLine($"Folder {file.Name} Created");

            return file;
        }

        private static void UpdateFile(string filePath, File folder, File file, string contentType)
        {
            File updatedFileMetadata = new File()
            {
                Name = file.Name,
                MimeType = "application/vnd.google-apps.spreadsheet"
            };

            Logger.Info($"Updateing File: {updatedFileMetadata.Name} in Folder: {folder.Name}");
            FilesResource.UpdateMediaUpload updateRequest;
            string fileId = file.Id;
            using (var stream = new FileStream(filePath, FileMode.Open))
            {
                updateRequest = service.Files.Update(updatedFileMetadata, fileId, stream, contentType);
                updateRequest.AddParents = folder.Id;
                updateRequest.Fields = "id,name";
                updateRequest.Upload();
            }
            var updatedFile = updateRequest.ResponseBody;
            if (!string.IsNullOrEmpty(updatedFile.Id))
            {
                Logger.Info($"File {updatedFile.Name} Updated");
            }
            Console.WriteLine($"File {updatedFile.Name} Updated");
        }

        #endregion

        #region General
        static void GetAppSettingsFile()
        {
            Logger.Info("Get App Settings File");
            var builder = new ConfigurationBuilder()
                 .SetBasePath(Directory.GetCurrentDirectory())
                 .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
            _iconfiguration = builder.Build();
        }

        private static string GetMimeType(string fileName)
        {
            string mimeType = "application/unknown";
            string ext = System.IO.Path.GetExtension(fileName).ToLower();
            Microsoft.Win32.RegistryKey regKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(ext);
            if (regKey != null && regKey.GetValue("Content Type") != null)
                mimeType = regKey.GetValue("Content Type").ToString();
            return mimeType;
        }

        private static string GetFilePath(string fileName)
        {
            string localFolderName = _iconfiguration.GetSection("LocalFolderName").Value;
            string startupPath = System.IO.Directory.GetCurrentDirectory();
            string directory = System.IO.Path.Combine(startupPath, localFolderName);
            if (!System.IO.Directory.Exists(directory))
            {
                System.IO.Directory.CreateDirectory(directory);
            }

            string filePath = System.IO.Path.Combine(directory, fileName);
            return filePath;
        }

        private static void SaveCurrentDateToFile()
        {
            Logger.Info($"Save Last Run Date to file ReportsLastRunDate.txt");

            var lastUsedDate = DateTime.Now.AbsoluteEnd().ToString();
            string fileName = $"ReportsLastRunDate.txt";
            string filePath = GetFilePath(fileName);
            System.IO.File.WriteAllText(filePath, lastUsedDate);
        }

        public static DateTime GetLastUsedDateFile()
        {
            Logger.Info($"Get Last Run Date from ReportsLastRunDate.txt file");

            var reportInitialDate = _iconfiguration.GetSection("ReportsInitialDate").Value;
            DateTime from = Convert.ToDateTime(reportInitialDate);

            string fileName = $"ReportsLastRunDate.txt";
            string filePath = GetFilePath(fileName);
            if (System.IO.File.Exists(filePath))
            {
                string[] lines = System.IO.File.ReadAllLines(filePath);
                foreach(var line in lines)
                {
                    if (!string.IsNullOrWhiteSpace(line) && DateTime.TryParse(line, out DateTime value))
                        from = value;
                }
            }
            return from.AbsoluteStart();
        }
        #endregion
    }
}