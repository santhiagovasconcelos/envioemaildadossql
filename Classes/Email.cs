using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using System.Text;
using System.Windows.Forms;

namespace ND_Wartsila.Classes
{
    class Email
    {
        String valorFormatado;

        public string FormatarValor(double valor)
        {
            valorFormatado = Convert.ToString(string.Format("{0:N}", valor));
            //valorFormatado = String.Format("{0:0,0.00}", valor);

            return valorFormatado;
        }

        private string AlterarNome(string nomeAntigo, string nomeNovo, string mesAno)
        {
            //TrimEnd() remove os espaços somente do fim da string.
            var nomeNovoCorrigido = nomeNovo.TrimEnd();//Output: " Texto a ser removidos seus espaços."

            //string oldFile = @"D:\oldfile.txt";
            string oldFile = "\\\\192.168.0.3\\entrada\\" + nomeAntigo;
            string newFile = "C:\\ND_WARTSILA\\" + nomeNovoCorrigido + " " + mesAno + ".pdf";
          
            File.Copy(oldFile, newFile);

            return newFile;

        }
        public void EnviarND(string emailPara, string emailLogin, string emailSenha, string caminhoClaroFixo, string caminhoClaroDados, string caminhoWorldnetDados, string data, string valorClaroDados, string valorClaroFixo, string valorWoldnetDados, string descricaoClaroFixo, string descricaoClaroDados, string descricaoWorldnetDados)
        {
            string assunto = "Geração de nota de Débito para Wartsila - Serviços de voz e dados - " + data + ".";
            //assuntoCompleto = assuntoInicio + nomeCompleto + assuntoFim;

            string valorWatsilaClaroFixo = FormatarValor(Convert.ToDouble(valorClaroFixo) /2);
            string valorWatsilaClaroDados = FormatarValor(Convert.ToDouble(valorClaroDados) / 2); 
            string valorWatsilaWorldnet = FormatarValor(Convert.ToDouble(valorWoldnetDados) / 2);

            string valorTotal = FormatarValor(Convert.ToDouble(valorWatsilaClaroFixo) + Convert.ToDouble(valorWatsilaClaroDados) + Convert.ToDouble(valorWatsilaWorldnet));

            string valorConvertidoClaroFixo = FormatarValor(Convert.ToDouble(valorClaroFixo));
            string valorConvertidoClaroDados = FormatarValor(Convert.ToDouble(valorClaroDados)); ;
            string valorConvertidoWoldnetDados = FormatarValor(Convert.ToDouble(valorWoldnetDados)); ;
           

            string primeiraParte = "<p style='font-family: verdana; font-size: 9pt;'>Prezado(a), </br></br>Em anexo seguem as faturas e planilha para acompanhamento e geração da nota de débito para a Wartsila, referente serviços de voz e dados do mês de " + data + ".";
            string segundaParte = "";
            string terceiraParte = "</br></br><b>Fatura telefone fixo:</b> R$" + valorConvertidoClaroFixo + ",</br><b>Custo Wartsila: </b>R$ "+ valorWatsilaClaroFixo;
            string quartaParte = "</br></br><b>Fatura dados Embratel:</b> R$" + valorConvertidoClaroDados + ",</br><b>Custo Wartsila: </b>R$ "+ valorWatsilaClaroDados;
            string quintaParte = "</br></br><b>Fatura dados Worldnet:</b> R$" + valorConvertidoWoldnetDados + ",</br><b>Custo Wartsila: </b>R$ "+ valorWatsilaWorldnet;
            string sextaParte = "</br></br><b>Total Wartsila:</b> R$" + valorTotal + ".";

            string setimaParte = "</br></br>Este e-mail foi enviado automaticamente pelo Sistema de envio de ND da ENERGETICA SUAPE II S.A.";
            string oitavaParte = "</br></br><a href='https://seudominio.com.br'>powered by Tecnologia da Informação | Suape Energia</a></p><hr></br><img width='200' src = 'https://seudominio.com.br/_imagens/logo-horizontal.png'/>";


            string corpoEmail = primeiraParte + segundaParte + terceiraParte + quartaParte + quintaParte + sextaParte + setimaParte + oitavaParte;

            //Attachment attClaroFixo = new Attachment(caminhoClaroFixo);
            //Attachment attClaroDados= new Attachment(caminhoClaroDados);
            //Attachment attWorldnetDados = new Attachment(caminhoWorldnetDados);  
            
            Attachment attClaroFixo = new Attachment(AlterarNome(caminhoClaroFixo, descricaoClaroFixo, data));
            Attachment attClaroDados= new Attachment(AlterarNome(caminhoClaroDados, descricaoClaroDados, data));
            Attachment attWorldnetDados = new Attachment(AlterarNome(caminhoWorldnetDados, descricaoWorldnetDados, data));

            Attachment attExcel = new Attachment("C:\\ND_WARTSILA\\Detalhamento.xlsx");

            //Attachment att = new Attachment(textDiretorio.Text + "\\" + boletoPdf);

            SmtpClient SmtpServer = new SmtpClient("outlook.office365.com");
            var mail = new MailMessage();
            mail.From = new MailAddress(emailLogin); //E-mail do remetente - o mesmo do login
            mail.To.Add(emailPara);
            mail.Subject = assunto;
            mail.IsBodyHtml = true;
            mail.Body = corpoEmail;
            SmtpServer.Port = 587;
            SmtpServer.UseDefaultCredentials = false;
            SmtpServer.Credentials = new System.Net.NetworkCredential(emailLogin, emailSenha);
            SmtpServer.EnableSsl = true;
            mail.Attachments.Add(attClaroFixo);
            mail.Attachments.Add(attClaroDados);
            mail.Attachments.Add(attWorldnetDados);
            mail.Attachments.Add(attExcel);

            try
            {
                SmtpServer.Send(mail);

            }
            catch (Exception ex)// deixando registro de erro no envio
            {
                MessageBox.Show("Erro: " + ex.ToString());

                //listLog.Items.Add("Erro: " + ex.ToString());
            }

        }


    }
}
