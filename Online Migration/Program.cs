using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;

namespace Online_Migration
{
    class Program
    {
        static void Main(string[] args)
        {
            string siteUrl = GetSite();
           // Console.WriteLine("Enter credentials for {0}", siteUrl);

            string userName = GetUserName();
            SecureString pwd = GetPassword();
        
            if (string.IsNullOrEmpty(userName) || (pwd == null))
                return;

            try
            {
                Console.WriteLine("Checking Default Authentication");
                ClientContext ctx = new ClientContext(siteUrl);
                ctx.AuthenticationMode = ClientAuthenticationMode.Default;
                ctx.Credentials = new SharePointOnlineCredentials(userName, pwd);
                Web web = ctx.Web;
                ctx.Load(web);
                ctx.ExecuteQuery();
                Console.WriteLine("the current web : " + web.Title);
                
                Console.WriteLine("Default Authentication successed");
               
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
              

            }
            try
            {
                Console.WriteLine("Checking Multi Authentication");
                var authenticationManager =new OfficeDevPnP.Core.AuthenticationManager();
                ClientContext context =authenticationManager.GetWebLoginClientContext(siteUrl, null);
                Web web = context.Web;        
                context.Load(web);           
                context.ExecuteQuery();
                 Console.WriteLine("the current web : " + web.Title);

                Console.WriteLine("Multi Athentication Authentication successed");
                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
             
            }
            Console.ReadLine();

        }

        public static Microsoft.SharePoint.Client.File UploadFile(ClientContext ctx, string Source, string filePath) {

             string filerelativeurl = string.Format("{0}/{1}", Source, System.IO.Path.GetFileName(filePath));
            //string filerelativeurl = Source;
            using (FileStream fs = new FileStream(filePath, FileMode.Open))
            {
                Microsoft.SharePoint.Client.File.SaveBinaryDirect(ctx, filerelativeurl, fs, true);
            }
            Microsoft.SharePoint.Client.File file = ctx.Web.GetFileByServerRelativeUrl(filerelativeurl);
            ctx.Load(file);
            ctx.ExecuteQuery();
            return file;


        }
        public static FieldLookupValue GetlookupValue(ClientContext ctx2, string keyword)
        {
          
            List list = ctx2.Web.Lists.GetByTitle("Application Master");  
            CamlQuery camlQuery = new CamlQuery();
            camlQuery.ViewXml = "<View><Query><Where><Eq><FieldRef Name='SystemCode' /><Value Type='Text'>"+ keyword + "</Value></Eq></Where></Query></View>";
            var itemCol = list.GetItems(camlQuery);
            ctx2.Load(itemCol);
            ctx2.ExecuteQuery();
            FieldLookupValue lookup = new FieldLookupValue();
            foreach (var item in itemCol)
            {
                 lookup.LookupId = item.Id;
            }
            return lookup;
        }
        public static Folder CreateFolder(Web web, string listTitle, string fullFolderUrl)
        {
            if (string.IsNullOrEmpty(fullFolderUrl))
                throw new ArgumentNullException("fullFolderUrl");
            var list = web.Lists.GetByTitle(listTitle);
            return CreateFolderInternal(web, list.RootFolder, fullFolderUrl);
        }

        private static Folder CreateFolderInternal(Web web, Folder parentFolder, string fullFolderUrl)
        {
            var folderUrls = fullFolderUrl.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
            string folderUrl = folderUrls[0];
            var curFolder = parentFolder.Folders.Add(folderUrl);
            web.Context.Load(curFolder);
            web.Context.ExecuteQuery();

            if (folderUrls.Length > 1)
            {
                var subFolderUrl = string.Join("/", folderUrls, 1, folderUrls.Length - 1);
                return CreateFolderInternal(web, curFolder, subFolderUrl);
            }
            return curFolder;
        }
        public static bool GetFolder(Web web, string fullFolderUrl)
        {
            try
            {
                if (string.IsNullOrEmpty(fullFolderUrl))
                    throw new ArgumentNullException("fullFolderUrl");

                if (!web.IsPropertyAvailable("ServerRelativeUrl"))
                {
                    web.Context.Load(web, w => w.ServerRelativeUrl);
                    web.Context.ExecuteQuery();
                }
                var folder = web.GetFolderByServerRelativeUrl(web.ServerRelativeUrl + fullFolderUrl);
                web.Context.Load(folder);
                web.Context.ExecuteQuery();
                return true;
            }
            catch(Exception ex)
            {
                return false;
            }
        }
        static SecureString GetPassword()
        {
            SecureString sStrPwd = new SecureString();
            try
            {
                Console.Write("Password: ");

                for (ConsoleKeyInfo keyInfo = Console.ReadKey(true); keyInfo.Key != ConsoleKey.Enter; keyInfo = Console.ReadKey(true))
                {
                    if (keyInfo.Key == ConsoleKey.Backspace)
                    {
                        if (sStrPwd.Length > 0)
                        {
                            sStrPwd.RemoveAt(sStrPwd.Length - 1);
                            Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                            Console.Write(" ");
                            Console.SetCursorPosition(Console.CursorLeft - 1, Console.CursorTop);
                        }
                    }
                    else if (keyInfo.Key != ConsoleKey.Enter)
                    {
                        Console.Write("*");
                        sStrPwd.AppendChar(keyInfo.KeyChar);
                    }

                }
                Console.WriteLine("");
            }
            catch (Exception e)
            {
                sStrPwd = null;
                Console.WriteLine(e.Message);
            }

            return sStrPwd;
        }

        static string GetUserName()
        {
            string strUserName = string.Empty;
            try
            {
                Console.Write("Username: ");
                strUserName = Console.ReadLine();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                strUserName = string.Empty;
            }
            return strUserName;
        }

        static string GetSite()
        {
            string siteUrl = string.Empty;
            try
            {
                Console.Write("Enter your Office365 site collection URL: ");
                siteUrl = Console.ReadLine();
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                siteUrl = string.Empty;
            }
            return siteUrl;
        }
      
        private static void FillTableWithData(string tableName, OleDbConnection conn, DataSet ds)
        {
            var query = string.Format("SELECT * FROM [{0}]", tableName);
            using (var da = new OleDbDataAdapter(query, conn))
            {
                var dt = new DataTable(tableName);
                da.Fill(dt);
                ds.Tables.Add(dt);
            }
        }

        private static IEnumerable<string> GetSheets(OleDbConnection conn)
        {
            var sheets = new List<string>();
            var dtSheet = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            if (dtSheet != null)
            {
                foreach (DataRow drSheet in dtSheet.Rows)
                {
                    string sheet = drSheet["TABLE_NAME"].ToString();
                    if (sheet.EndsWith("$") || sheet.StartsWith("'") && sheet.EndsWith("$'"))
                    {
                        sheets.Add(sheet);
                    }
                }
            }

            return sheets;
        }

        private static string GetExcelConnectionString(string fullPathToExcelFile)
        {
            var props = new Dictionary<string, string>();
            props["Provider"] = "Microsoft.ACE.OLEDB.12.0";
            props["Data Source"] = fullPathToExcelFile;
            props["Extended Properties"] = "\"Excel 12.0;HDR=Yes;IMEX=1\"";
           // props["HDR"] = "Yes";
            var sb = new StringBuilder();
            foreach (var prop in props)
            {
                sb.Append(prop.Key);
                sb.Append('=');
                sb.Append(prop.Value);
                sb.Append(';');
            }

            return sb.ToString();
        }
    }
}
