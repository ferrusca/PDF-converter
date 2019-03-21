using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text;
using Microsoft.Office.Word.Server.Conversions;
using Microsoft.SharePoint;
namespace Test
{
    class Converter
    {
        public static object HttpContext { get; private set; }

        static void Main(string[] args)
        {
            using (SPSite spSite = new SPSite("http://proyecto:42091/"))
            {
                var deptos = new string[4] {"Depto1", "Depto2", "Depto3", "Depto4"};
                String mailBody = "";
                foreach (string depto in deptos) { 
                    var library = spSite.RootWeb.Lists.TryGetList(depto);
                    SPQuery query = new SPQuery();
                    query.Folder = library.RootFolder;
                    //Include all subfolders so include Recursive Scope.
                    query.ViewXml = @"<View Scope='Recursive'>
                        <Query>
                            <Where>
                                <Or>
                                    <Contains>
                                        <FieldRef Name='File_x0020_Type'/>
                                        <Value Type='Text'>doc</Value>
                                    </Contains>
                                    <Contains>
                                        <FieldRef Name='File_x0020_Type'/>
                                        <Value Type='Text'>docx</Value>
                                    </Contains>
                                </Or>
                            </Where>
                        </Query>
                    </View>";
                    //Obtaining files from query result
                    SPListItemCollection listItems = library.GetItems(query);
                    if (listItems.Count > 0)
                    {
                        mailBody += "<b>Archivos de " + depto + "</b><br><br><ul>";
                        ConversionJobSettings jobSettings = new ConversionJobSettings();
                        jobSettings.OutputFormat = SaveFormat.PDF;
                        SyncConverter pdfConversion = new SyncConverter("Word Automation Services", jobSettings);
                        pdfConversion.UserToken = spSite.UserToken;
                        foreach (SPListItem li in listItems)
                        {
                            string fileSource = (string)li[SPBuiltInFieldId.EncodedAbsUrl];
                            string fileDest = fileSource.Replace("docx", "pdf");
                            fileDest = fileDest.Replace("doc", "pdf");
                            fileDest = fileDest.Replace(depto, "Documentacion");
                            mailBody += "<li><b>Origen: </b>" + fileSource + "</li>";
                            mailBody += "<li><b>Destino: </b>" + fileDest + "</li><br>";
                            Console.Write("Origen: " + fileSource + "\n");
                            Console.Write("Dest: " + fileDest + "\n\n");
                            pdfConversion.Convert(fileSource, fileDest);
                        }
                        mailBody += "</ul><br><br>";
                    } //foreach depto
                    try
                    {
                        MailMessage mail = new MailMessage();
                        SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
                        mail.IsBodyHtml = true;
                        mail.From = new MailAddress("foobar.dgtic@gmail.com");
                        mail.To.Add("foobar.dgtic@gmail.com");
                        mail.Subject = "Lista de Archivos";
                        mail.Body = mailBody;
                        SmtpServer.Port = 587;
                        SmtpServer.Credentials = new System.Net.NetworkCredential("foobar.dgtic@gmail.com", "******");
                        SmtpServer.EnableSsl = true;
                        SmtpServer.Send(mail);
                        Console.Write("Mail enviado");
                    }
                    catch (Exception ex)
                    {
                        Console.Write(ex.ToString());
                    }
                }
            }
        }
    }
}