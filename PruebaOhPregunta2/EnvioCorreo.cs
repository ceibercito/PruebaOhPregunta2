using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace PruebaOhPregunta2
{
    internal class EnvioCorreo
    {
        public string envioCorreo(string to, string asunto, string body, string email, string clave)
        {

            string mensaje = "Error al enviar este correo, Por qfavor verifique los datos o intente mas tarde";
            string from = email;
            string displayName = "CEIBER CONRADO GARIBAY CHOQUE";
            string path = Directory.GetCurrentDirectory() + "\\Reporte.xlsx ";

            try
            {

                MailMessage correo = new MailMessage();
                correo.Attachments.Add(new Attachment(path));
                correo.From = new MailAddress(from, displayName);
                correo.To.Add(to);

                correo.Subject = asunto;
                correo.Body = body;
                correo.IsBodyHtml = true;

                SmtpClient cliente = new SmtpClient("smtp.office365.com", 587);
                cliente.Credentials = new NetworkCredential(from, clave);
                cliente.EnableSsl = true;

                cliente.Send(correo);
                mensaje = "¡Correo enviado exitosamente!";
                cliente.Dispose();
            }
            catch(Exception ex)
            {
                mensaje = ex.Message + ". Por favor verifica tu conexion a internet y que tus datos sean correctos e intenta nuevamente";
            }
            return mensaje;
        }
    }
}
