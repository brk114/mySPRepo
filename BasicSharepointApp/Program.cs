using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint;
using System.Security;
using System.IO;
using System.Net;

namespace BasicSharepointApp
{
    class Program
    {
        static SharePointOnlineCredentials GetCredentials()
        {
            var pwd = "Mugils$6";
            var secureString = new SecureString();
            Array.ForEach(pwd.ToCharArray(), x => secureString.AppendChar(x));

            var spCredentials = new SharePointOnlineCredentials("ravi@ravib.onmicrosoft.com", secureString);

            return spCredentials;
        }

        static void Main(string[] args)
        {
            
        }

        static void Code() {
            try
            {
                using (var context = new ClientContext("https://ravib.sharepoint.com"))
                {
                    //Add Credentials
                    context.Credentials = GetCredentials();

                    Web web = context.Web;

                    //Loading Site Groups
                    #region Site groups and User Info In Each group

                    //context.Load(web.SiteGroups, x => x.Include(sg => sg.Title, sg => sg.Id, sg => sg.Users));
                    //// Execute the query to server.
                    //context.ExecuteQuery();

                    ////Console.WriteLine("Web tittle : " + web.Title);
                    //Array.ForEach(web.SiteGroups.ToArray(), x =>
                    //{
                    //    if (x != null)
                    //    {
                    //        Console.WriteLine(x.Id + " : " + x.Title);
                    //        if (x.Users != null)
                    //            Array.ForEach(x.Users.ToArray(), y =>
                    //            {
                    //                if (y.UserId != null && y.UserId.NameId != null && y.Email != null)
                    //                    Console.Write("\n ===>UID: " + y.UserId.NameId + ", Email:" + y.Email);
                    //                else Console.WriteLine("====================> No Users");
                    //            });
                    //    }
                    //    else Console.WriteLine("====================> No Content");
                    //});

                    #endregion

                    #region Create New List
                    //ListCreationInformation creationInfo = new ListCreationInformation();
                    //creationInfo.Title = "My Programatic Custom List";
                    //creationInfo.TemplateType = (int)ListTemplateType.GenericList;
                    //List lst = web.Lists.Add(creationInfo);
                    //lst.Description = "New Description";

                    //lst.Update();
                    //context.ExecuteQuery(); 
                    #endregion

                    #region Delete List
                    //List list = web.Lists.GetByTitle("My Programatic List");
                    //list.DeleteObject();
                    //context.ExecuteQuery();
                    #endregion

                    #region All List Id Tittle

                    /*
                    context.Load(web, w => w.Lists);
                    context.ExecuteQuery();

                    // Retrieve all lists from the server, and put the return value in another 
                    // collection instead of the web.Lists. 
                    IEnumerable<List> result = context.LoadQuery(web.Lists.Include(list => list.Title, list => list.Id));

                    // Execute query. 
                    context.ExecuteQuery();

                    // Enumerate the result. 
                    foreach (List list in web.Lists)
                    {
                        Console.WriteLine(list.Id + ", " + list.Title);
                    }
                    */

                    #endregion

                    #region Display all list Items
                    //context.Load(web, w => w.Lists);

                    //List list = context.Web.Lists.GetByTitle("product");
                    //CamlQuery query = CamlQuery.CreateAllItemsQuery(100);
                    //ListItemCollection items = list.GetItems(query);

                    //context.Load(items);
                    //context.ExecuteQuery();
                    //Console.WriteLine("ID   Title   Product Name \tPrice");
                    //foreach (ListItem listItem in items)
                    //{
                    //    //Console.WriteLine(listItem["Title"]);
                    //    Console.WriteLine(listItem["ID"] + "  "+ listItem["Title"] + "  " + listItem["Product_x0020_Name"] + " \t\t" + listItem["Price"]);
                    //}

                    #endregion

                    #region All Columns of a List
                    //context.Load(web, w => w.Lists);

                    //List list = context.Web.Lists.GetByTitle("product");
                    //context.Load(list.Fields);

                    //// We must call ExecuteQuery before enumerate list.Fields. 
                    //context.ExecuteQuery();

                    //foreach (Field field in list.Fields)
                    //{
                    //    Console.WriteLine(field.InternalName);
                    //}
                    #endregion

                    #region Upload file to Document Library

                    //string fileName = @"E:\abc.txt";
                    //var newFile = new FileCreationInformation
                    //{
                    //    Content = System.IO.File.ReadAllBytes(fileName),
                    //    Url = Path.GetFileName(fileName)
                    //};
                    //var docs = web.Lists.GetByTitle("Documents");
                    //Microsoft.SharePoint.Client.File uploadFile = docs.RootFolder.Files.Add(newFile);
                    //context.ExecuteQuery();
                    #endregion

                    #region Add a User toa user group

                    //GroupCollection userGroups = context.Web.SiteGroups;
                    //Group group = userGroups.GetByName("EmployeeGroup");
                    //var userCreationInfo = new UserCreationInformation()
                    //{
                    //    Email = "Karthikb@ravib.onmicrosoft.com",
                    //    LoginName = "Karthikb",
                    //    Title = "Karthik B"
                    //}; 
                    #endregion

                    #region Create a SubSite

                    //WebCreationInformation subSite = new WebCreationInformation()
                    //{
                    //    Url = "Test",
                    //    Title = "Test Web",
                    //    WebTemplate = "BLOG#0"
                    //};
                    //Web newWeb = context.Web.Webs.Add(subSite);
                    //context.Load(newWeb, w => w.Title);
                    //context.ExecuteQuery();
                    //Console.WriteLine("New Tittle: " + newWeb.Title);
                    #endregion

                    #region All Sub Sites

                    //context.Load(web, w => w.Webs);
                    //context.ExecuteQuery();
                    //// Retrieve the new web information.
                    //WebCollection subSites = web.Webs;
                    //Array.ForEach(subSites.ToArray(), x => {
                    //    context.Load(x, w=>w.Title);
                    //});

                    //context.ExecuteQuery();

                    //Array.ForEach(subSites.ToArray(), x => {
                    //    Console.WriteLine(x.Title);
                    //});
                    #endregion

                    #region Download a file

                    //var list = context.Web.Lists.GetByTitle("Documents");
                    //var listItem = list.GetItemById(16);
                    //context.Load(list);
                    //context.Load(listItem, li => li.File);
                    //context.ExecuteQuery();

                    //var fileRef = listItem.File.ServerRelativeUrl;
                    //var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(context, fileRef);
                    //var fileName = Path.Combine("E:/", "newDolededFile.txt");
                    //using (var fileStream = System.IO.File.Create(fileName))
                    //{
                    //    fileInfo.Stream.CopyTo(fileStream);
                    //}
                    //Console.WriteLine("File Downloaded");
                    #endregion

                    #region Delete a file
                    //List docs = web.Lists.GetByTitle("Documents");

                    //Microsoft.SharePoint.Client.File file = web.GetFileByServerRelativeUrl("/Shared Documents/abc.txt");
                    //context.Load(file);
                    //file.DeleteObject();
                    //context.ExecuteQuery(); // Delete file here but throw Exception                
                    //Console.WriteLine("File deleted");
                    #endregion

                    #region Creating Custom content Type
                    ////// Specifies properties that are used as parameters to initialize a new content type.
                    //ContentTypeCreationInformation contentTypeCreation = new ContentTypeCreationInformation()
                    //{
                    //    Name = "Customer",
                    //    Description = "Custom Content Type created using CSOM",
                    //    Group = "List Content Types"
                    //};

                    ////// Get the content type collection for the website
                    //ContentTypeCollection contentTypeColl = context.Web.ContentTypes;

                    ////// Add the new content type to the collection
                    //ContentType contentType = contentTypeColl.Add(contentTypeCreation);
                    //context.Load(contentType);
                    //context.ExecuteQuery();

                    ////// Display that the content type is created.
                    //Console.WriteLine(contentType.Name + " content type is created successfully");
                    #endregion
                    #region Get All Custom TYpes By ID
                    //ContentTypeCollection contentTypeColl = context.Web.ContentTypes;

                    //context.Load(contentTypeColl);
                    //context.ExecuteQuery();


                    //Array.ForEach(contentTypeColl.OrderBy(x => x.Name).ToArray(), cType =>
                    //{
                    //    Console.WriteLine("ID " + cType.Id + "\t\t Name " + cType.Name);
                    //}); 
                    #endregion

                    #region Adding Columns To the Custom Content Type
                    //ContentTypeCollection contentTypeColl = context.Web.ContentTypes;
                    //ContentType contentType = contentTypeColl.GetById("0x0100B16668BA0BB3014B9D6F62C18B6CEF8B");

                    //string[] fields = new string[] { "CustomerAdress", "CustomerName", "CustomerType" };
                    //web.Fields.AddFieldAsXml("<Field DisplayName='Customer Address' Name='" + fields[0] + "'    Group='MyFieldTypes' Type='Text' />", false, AddFieldOptions.AddFieldInternalNameHint);
                    //web.Fields.AddFieldAsXml("<Field DisplayName='Customer Name' Name='" + fields[1] + "'       Group='MyFieldTypes' Type='Text' />", false, AddFieldOptions.AddFieldInternalNameHint);
                    //web.Fields.AddFieldAsXml("<Field DisplayName='Customer Type' Name='" + fields[2] + "'       Group='MyFieldTypes' Type='Text' />", false, AddFieldOptions.AddFieldInternalNameHint);
                    //context.ExecuteQuery();

                    //Array.ForEach(fields, field =>
                    //{
                    //    contentType.FieldLinks.Add(new FieldLinkCreationInformation()
                    //    {
                    //        Field = web.Fields.GetByInternalNameOrTitle(field),
                    //    });
                    //});

                    //contentType.Update(true);
                    //context.ExecuteQuery();
                    #endregion

                    #region Creating Cutom List With Custom Content type
                    //ListCreationInformation creationInfo = new ListCreationInformation();
                    //creationInfo.Title = "MyListWithCustomType";
                    //creationInfo.TemplateType = (int)ListTemplateType.GenericList;
                    //List list = web.Lists.Add(creationInfo);
                    //list.Description = "Programatic custom list with custom content type";

                    //list.Update();
                    //context.ExecuteQuery();

                    ////Adding cutom type Customer
                    //ContentType contentType = web.ContentTypes.GetById("0x0100B16668BA0BB3014B9D6F62C18B6CEF8B");

                    ////// Add the content type to the custom list
                    //list.ContentTypes.AddExistingContentType(contentType);
                    //context.ExecuteQuery();
                    #endregion

                    #region Get all List level Contype IDs

                    //List list = web.Lists.GetByTitle("MyListWithCustomType");

                    //context.Load(list.ContentTypes);
                    //context.ExecuteQuery();
                    //Array.ForEach(list.ContentTypes.ToArray(), x =>
                    //{
                    //    Console.WriteLine("Name: " + x.Name +"; ID: " + x.Id);
                    //});.

                    #endregion

                    #region Add Items of custom content type to list
                    //List list = web.Lists.GetByTitle("MyListWithCustomType");
                    //context.Load(list.ContentTypes);
                    //context.ExecuteQuery();

                    ////Adding Customer custom type values
                    //ListItem li = list.AddItem(new ListItemCreationInformation());
                    //li["CustomerAdress"] = "Hederabad";
                    //li["CustomerName"] = "Karthik";
                    //li["CustomerType"] = "Gold";
                    //li["Title"] = "Test Record 3";
                    //li["ContentTypeId"] = list.ContentTypes.SingleOrDefault(x => x.Name.Equals("Customer")).Id;

                    //li.Update();
                    //context.ExecuteQuery();
                    #endregion

                    #region Working with CAML queries

                    //List list = web.Lists.GetById(new Guid("2ba45a86-71fe-4548-ab08-3cb3db28e7bb"));
                    //CamlQuery query = new CamlQuery()
                    //{
                    //    ViewXml = "<View><Query><Where><Eq><FieldRef Name='CustomerType' /><Value Type='Text'>Gold</Value></Eq></Where></Query><ViewFields><FieldRef Name='ID' /><FieldRef Name='CustomerAdress' /><FieldRef Name='CustomerName' /><FieldRef Name='CustomerType' /><FieldRef Name='ContentType' /><FieldRef Name='Title' /></ViewFields><QueryOptions /></View>"
                    //    //ViewXml = "<View><Query><Where><And><Eq><FieldRef Name='CustomerType' /><Value Type='Text'>Gold</Value></Eq><And><Eq><FieldRef Name='CustomerAdress' /><Value Type='Text'>Chennai</Value></Eq><Gt><FieldRef Name='ID' /><Value Type='Counter'>2</Value></Gt></And></And></Where></Query><ViewFields><FieldRef Name='ID' /><FieldRef Name='CustomerAdress' /><FieldRef Name='CustomerName' /><FieldRef Name='CustomerType' /><FieldRef Name='Title' /><FieldRef Name='ContentType' /></ViewFields><QueryOptions /></View>"
                    //};
                    //ListItemCollection result = list.GetItems(query);
                    //context.Load(result);
                    ////context.Load(list.Fields);
                    //context.ExecuteQuery();
                    //string[] cols = new string[] { "ID", "CustomerAdress", "CustomerName", "CustomerType", "ContentType" };

                    //Array.ForEach(cols, x => { Console.Write("\t" + x); });
                    //Array.ForEach(result.ToArray(), x =>
                    //{
                    //    Console.WriteLine();
                    //    Array.ForEach(x.FieldValues.Keys.ToArray(), key =>
                    //    {
                    //        if (cols.Contains(key))
                    //            Console.Write("\t" + x.FieldValues[key]);
                    //    });
                    //});


                    #endregion


                    #region insert Item to list

                    List list = web.Lists.GetByTitle("TestList");
                    context.Load(list);
                    context.ExecuteQuery();

                    //Adding Customer custom type values
                    ListItem li = null;
                    for (int i = 5001; i < 5050; i++)
                    {
                        li = list.AddItem(new ListItemCreationInformation());
                        li["Title"] = "Item_" + i;
                        //li["ContentTypeId"] = list.ContentTypes.SingleOrDefault(x => x.Name.Equals("Item")).Id;

                        li.Update();

                        Console.WriteLine("added Item " + i);
                        if (i % 10 == 0 || i > 5040)
                            context.ExecuteQuery();
                    }

                    #endregion
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message + "===> Message Exception");
            }

        }
    }
}
