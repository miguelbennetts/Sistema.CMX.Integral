using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace CMX360.Comunes.Clases
{
    public class Alerta
    {
        public string Asunto { get; set; }
        public string Contenido { get; set; }
        public string[] Destinatarios { get; set; }
        public string[] DestinatariosCC { get; set; }
        public string[] DestinatariosBcc { get; set; }
        public string MensajeSMS { get; set; }
        public string[] TelefonosSMS { get; set; }
        public List<ArchivoAdjunto> ArchivosAdjuntos { get; set; }

        public string MandaCorreo()
        {
            using (SmtpClient smtp = new SmtpClient(ConfigurationManager.AppSettings["smtp"]))
            {
                using (MailMessage correo = new MailMessage())
                {
                    correo.From = new MailAddress(ConfigurationManager.AppSettings.Get("CorreoContacto"),"Merezco Amarme");

                    if (this.Destinatarios != null)
                    {
                        foreach (string destinatario in this.Destinatarios)
                        {
                            correo.To.Add(destinatario);
                        }
                    }

                    if (this.DestinatariosCC != null)
                    {
                        foreach (string destinatario in this.DestinatariosCC)
                        {
                            correo.CC.Add(destinatario);
                        }
                    }

                    if (this.DestinatariosBcc != null)
                    {
                        foreach (string destinatario in this.DestinatariosBcc)
                        {
                            correo.Bcc.Add(destinatario);
                        }
                    }

                    if (this.ArchivosAdjuntos != null)
                    {

                        foreach (ArchivoAdjunto archivo in this.ArchivosAdjuntos)
                        {
                            correo.Attachments.Add(new Attachment(new MemoryStream(archivo.Archivo), archivo.Nombre));
                        }
                    }

                    //string logo = Path.Combine(HttpRuntime.AppDomainAppPath, $@"Content\images\{ConfigurationManager.AppSettings.Get("LogoCorreo")}");
                    //string nombrelogo = ConfigurationManager.AppSettings.Get("LogoCorreo");

                    //Attachment imgLogo = new Attachment(logo);
                    //imgLogo.ContentId = nombrelogo;
                    //correo.Attachments.Add(imgLogo);
                    smtp.DeliveryMethod = SmtpDeliveryMethod.Network;
                    correo.Subject = this.Asunto;
                    correo.IsBodyHtml = true;
                    correo.Body = Contenido;// GetPlantillaAlerta(this, nombrelogo);
                    correo.Priority = MailPriority.Normal;

                    try
                    {
                        smtp.Send(correo);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    return "Ok";
                }
            }
        }

        private string GetPlantillaAlerta(Alerta alerta, string logo)
        {
            StringBuilder sbCuerpo = new StringBuilder();
            string ruta = Path.Combine(HttpRuntime.AppDomainAppPath, @"Content\Plantillas\Alerta.html");
            using (StreamReader reader = new StreamReader(ruta))
            {
                sbCuerpo.Append(reader.ReadToEnd());
            }

            sbCuerpo = sbCuerpo.Replace("#LOGO#", string.Format("<img class='img-responsive' width='150px'  src=\'cid:{0}\'>", logo));

            sbCuerpo = sbCuerpo.Replace("#CONTENT#", alerta.Contenido);
            sbCuerpo = sbCuerpo.Replace("#FOOTER#", "Mensaje enviado automáticamente");
            sbCuerpo = sbCuerpo.Replace("#COPYRIGTH#", ConfigurationManager.AppSettings.Get("NombreCompania") + DateTime.Today.Year.ToString());

            return sbCuerpo.ToString();
        }
    }
}
