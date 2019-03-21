## Código Fuente
Para la realización del programa, importamos dos bibliotecas *dll*, en este caso __*Microsoft.SharePoint*__  y __*Microsoft.Office.Word*__. 

La clase principal obtiene el sitio en el cual está corriendo __*Sharepoint*__, y utilizando un arreglo que contiene el nombre de cada biblioteca, iteraremos de forma recursiva sobre cada una de ellas mediante una *query* en formato __XML__: 
```cs
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
```
Una vez obtenidos todos los items *.doc* o *.docx* de cada biblioteca, el proceso que realizaremos será el de conversion a pdf __*On demand*__ (se muestra resumido en el siguiente código):
```cs
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
    pdfConversion.Convert(fileSource, fileDest);
}
```                     
Como se puede ver, se inicializa un objeto para los *settings* de la conversion a realizar, especificando que el formato de salida será __PDF__. Una vez con estos ajustes, se invoca a la aplicación de servicio __Word Automation Services__, la cual nos ayudará a realizar dicha tarea de conversion.
Se obtiene la ruta de origen de cada archivo, y se genera una similar que además contiene el directorio *Documentacion*, y una vez definidos el origen y destino, se inicia la conversión de forma asíncrona, lo cual es logrado gracias al método *convert()* de la clase __*SyncConverter*__, la cual a diferencia de *ConvertionJob*, realiza la conversión del archivo inmediatamente, sin tener que poner en la cola a cada uno de los archivos que se desean convertir.

Finalmente se envía un mensaje al administrador, indicando los archivos que fueron copiados y de qué departamento. El código que lo realiza: 
```cs
MailMessage mail = new MailMessage();
SmtpClient SmtpServer = new SmtpClient("smtp.gmail.com");
mail.IsBodyHtml = true;
mail.From = new MailAddress("converter.dgtic@gmail.com");
mail.To.Add("foobar.dgtic@gmail.com");
mail.Subject = "Lista de Archivos";
mail.Body = mailBody;
SmtpServer.Port = 587;
SmtpServer.Credentials = new System.Net.NetworkCredential("foobar.dgtic@gmail.com", "hola123,");
SmtpServer.EnableSsl = true;
SmtpServer.Send(mail);
Console.Write("Mail enviado");
```

De esta forma, el código completo queda de la siguiente manera:
```cs
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
                        mail.From = new MailAddress("sender.dgtic@gmail.com");
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
```