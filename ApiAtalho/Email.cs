using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text;

namespace ApiAtalho
{
    public class Email
    {
        public void EnviarEmail(string[] linhas, string apontamento)
        {
            var body = MontaCorpo(linhas, apontamento);

            EnviaEmail(body, linhas[7]);
        }

        private string MontaCorpo(string[] linhas, string apontamento)
        {
            var body = File.ReadAllText($@"{AppDomain.CurrentDomain.BaseDirectory}modelo.html");
            body = body.Replace("{{apontamento}}", apontamento)
                        .Replace("{{dataSol}}", linhas[0])
                        .Replace("{{analista}}", linhas[1])
                        .Replace("{{tipoAtendimento}}", linhas[2])
                        .Replace("{{solicitante}}", linhas[3])
                        .Replace("{{area}}", linhas[4])
                        .Replace("{{dataInicio}}", linhas[5])
                        .Replace("{{dataFim}}", linhas[6])
                        .Replace("{{assunto}}", linhas[7])
                        .Replace("{{descricao}}", linhas[8])
                        .Replace("{{tempoGasto}}", linhas[9]);

            return body;
        }

        private void EnviaEmail(string body, string assunto)
        {
            var email = new MailMessage("josiel.alves@smn.com.br", "suporte.momentum@smn.com.br")
            {
                Subject = assunto,
                Body = body,
                IsBodyHtml = true,
                Priority = MailPriority.Normal,
                SubjectEncoding = Encoding.GetEncoding("ISO-8859-1"),
                BodyEncoding = Encoding.GetEncoding("ISO-8859-1")
            };

            email.Bcc.Add("josiel.alves@smn.com.br");

            try
            {
                using (var smtp = new SmtpClient("smtp-mail.outlook.com", 587) { Credentials = new NetworkCredential("josiel.alves@smn.com.br", "SMN@2017#o365") })
                {
                    smtp.EnableSsl = true;
                    smtp.Send(email);
                }
                Console.WriteLine("E-mail enviado com sucesso!");
            }
            catch (Exception erro)
            {
                Console.WriteLine("Erro ao enviar E-mail \n\n" + erro);
            }
        }
    }
}
