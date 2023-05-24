using Microsoft.Extensions.DependencyInjection;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Pdf.Security;
using Soneta.Business;
using Soneta.Business.Db;
using Soneta.Business.UI;
using Soneta.Core;
using Soneta.Deklaracje;
using Soneta.Deklaracje.PIT;
using Soneta.Deklaracje.ZUS;
using Soneta.Kadry;
using Soneta.Ksiega;
using Soneta.Tools;
using Soneta.Types;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;
using static Alex_WysylkaPIT.Program.Params;
//using Attachment = System.Net.Mail.Attachment;

[assembly: Worker(typeof(Alex_WysylkaPIT.Program), typeof(Pracownicy))]

namespace Alex_WysylkaPIT
{
    class Program
    {
        public class Params : ContextBase
        {
            private Pracownik[] wybraneOsoby;
            private bool PIT_bool;
            private bool IMIR_bool;


            public Params(Context context) : base(context) { }

            [Priority(1)]
            [Caption("Lista osób")]
            [Required]
            public Pracownik[] Pracownicy
            {
                get
                {
                    return this.wybraneOsoby;
                }
                set
                {
                    this.wybraneOsoby = value;
                    OnChanged(EventArgs.Empty);
                }
            }

            [Priority(10)]
            [Caption("Wyślij PIT")]
            public bool PIT_boolean
            {
                get
                {
                    return this.PIT_bool;
                }
                set
                {
                    this.PIT_bool = value;
                    OnChanged(EventArgs.Empty);
                }
            }

            [Priority(20)]
            [Caption("Wyślij IMIR")]
            public bool IMIR_boolean
            {
                get
                {
                    return this.IMIR_bool;
                }
                set
                {
                    this.IMIR_bool = value;
                    OnChanged(EventArgs.Empty);
                }
            }
        }
        private Params param;

        [Context]
        public Params Parametry
        {
            get { return param; }
            set { param = value; }
        }

        //[Context(CreateOnly = true)]
        //public XtraReportSerialization.Pit1129.Params Parms { get; set; }
        //[Context]
        //public XtraReportSerialization.Pit1129.Params Parms { get; set; }

        [Context]
        public Context Context { get; set; }

        [Action("PIT | IMIR do pracownika",
        Target = ActionTarget.ToolbarWithText,
        Priority = 1001,
        Icon = ActionIcon.ArrowDown,
        Mode = ActionMode.SingleSession)]
        public MessageBoxInformation Import()
        {
            int wyslano = 0;
            int brak_email = 0;
            int brak_deklaracji = 0;
            int blad_report_result = 0;
            int wyslanych_wczesniej = 0;
            List<string> listaPlikowDoUsuniecia = new List<string>();

            ServicePointManager.ServerCertificateValidationCallback += (sender, cert, chain, sslPolicyErrors) => true;

            foreach (Pracownik pracownik in param.Pracownicy)
            {
                // ### Walidacja email
                if (pracownik.Kontakt.EMAIL == "")
                {
                    brak_email++;
                    continue;
                }

                List<System.Net.Mail.Attachment> listaAtt = new List<System.Net.Mail.Attachment>();
                BusinessModule bm = BusinessModule.GetInstance(Context.Session);
                KadryModule km = KadryModule.GetInstance(Context.Session);
                KsiegaModule ksm = KsiegaModule.GetInstance(Context.Session);
                DeklaracjeModule dm = DeklaracjeModule.GetInstance(Context.Session);

                string numer_pit = "";
                string numer_imir = "";

                bool potwierdzenie_pit = false;
                bool potwierdzenie_imir = false;  

                if (param.PIT_boolean)
                {
                    Context cx = Context.Empty.Clone(Context.Session);
                    var attNamePDF = pracownik.ImięNazwisko.ToUpper() + " - PIT.pdf";
                    var dirName = Environment.GetEnvironmentVariable("TMP"); //"C:\\TEMP\\";
                    if (!System.IO.Directory.Exists(dirName))
                        System.IO.Directory.CreateDirectory(dirName);

                    var pathFileNamePDF = System.IO.Path.Combine(dirName, attNamePDF);

                    IReportService rs;
                    rs = cx.Session.GetRequiredService<IReportService>();
                    if (rs == null)
                    {
                        blad_report_result++;
                        continue;
                    }

                    int rok_pitu = Date.Today.AddYears(-1).Year;

                    View view_PIT = dm.Deklaracje.CreateView();
                    view_PIT.Condition &= new FieldCondition.Equal("Definicja.Symbol", "PIT11");
                    view_PIT.Condition &= new FieldCondition.Equal("Podmiot", pracownik);
                    view_PIT.Condition &= new FieldCondition.Equal("Okres.From", new Date(rok_pitu, 1, 1));
                    view_PIT.Sort = "ID";
                    if (view_PIT.Count == 0)
                    {
                        brak_deklaracji++;
                        continue;
                    }
                    PIT11_29 deklaracjaPIT = (PIT11_29)view_PIT.GetLast();
                    if ((bool)deklaracjaPIT.Features["Wysłano e-mail"])
                    {
                        wyslanych_wczesniej++;
                    }
                    else
                    {

                        numer_pit = deklaracjaPIT.Numer.NumerPelny;
                        cx.Set(deklaracjaPIT);

                        PIT11_29[] deklaracje = { deklaracjaPIT };
                        cx.Set(deklaracje);
                        cx.Set(new SelectedCounter(cx));

                        ActualDate aDate = new ActualDate(cx);
                        cx.Set(aDate);

                        // ### PARAMETRY PIT11 ###
                        cx.Set(new XtraReportSerialization.Pit1129.Params(cx) { PITR = false });

                        var report = new Soneta.Business.UI.ReportResult()
                        {
                            Context = cx,
                            TemplateFileSource = Soneta.Business.UI.AspxSource.Storage,
                            Format = ReportResultFormat.PDF,
                            //DataType = typeof(Deklaracja),
                            DataType = typeof(PIT11_29),
                            TemplateFileName = @"XtraReports/Wzorce użytkownika/alex_pit_29.repx",
                        };


                        using (Stream stream2 = rs.GenerateReport(report))
                        {
                            using (var file = System.IO.File.Create(pathFileNamePDF))
                            {
                                CoreTools.StreamCopy(stream2, file);
                                file.Flush();
                            }


                        }

                        // ### NADAWANIE HASŁA ###
                        PdfDocument document = PdfReader.Open(pathFileNamePDF);
                        PdfSecuritySettings securitySettings = document.SecuritySettings;
                        securitySettings.UserPassword = pracownik.PESEL.Substring(6);
                        document.Save(pathFileNamePDF);

                        listaAtt.Add(new System.Net.Mail.Attachment(pathFileNamePDF));

                        potwierdzenie_pit = true;
                    }
                }

                if (param.IMIR_boolean)
                {
                    Context cx = Context.Empty.Clone(Context.Session);
                    var attNamePDF = pracownik.ImięNazwisko.ToUpper() + " - IMIR.pdf";
                    var dirName = Environment.GetEnvironmentVariable("TMP");
                    if (!System.IO.Directory.Exists(dirName))
                        System.IO.Directory.CreateDirectory(dirName);

                    var pathFileNamePDF = System.IO.Path.Combine(dirName, attNamePDF);

                    IReportService rs;
                    rs = cx.Session.GetRequiredService<IReportService>();
                    if (rs == null)
                    {
                        blad_report_result++;
                        continue;
                    }

                    int rok_imiru = Date.Today.AddYears(-1).Year;

                    View view_IMIR = dm.Deklaracje.CreateView();
                    view_IMIR.Condition &= new FieldCondition.Equal("Definicja.Symbol", "IMIR");
                    view_IMIR.Condition &= new FieldCondition.Equal("Podmiot", pracownik);
                    view_IMIR.Condition &= new FieldCondition.Equal("Okres.From", new Date(rok_imiru, 1, 1));
                    view_IMIR.Sort = "ID";
                    if (view_IMIR.Count == 0)
                    {
                        brak_deklaracji++;
                        continue;
                    }
                    RMUA deklaracjaIMIR = (RMUA)view_IMIR.GetLast();
                    if ((bool)deklaracjaIMIR.Features["Wysłano e-mail"])
                    {
                        wyslanych_wczesniej++;
                    }
                    else
                    {
                        numer_imir = deklaracjaIMIR.Numer.NumerPelny;

                        cx.Set(deklaracjaIMIR);

                        RMUA[] deklaracje = { deklaracjaIMIR };
                        cx.Set(deklaracje);
                        cx.Set(new CurrentObject(cx));
                        cx.Set(new SelectedCounter(cx));

                        ActualDate aDate = new ActualDate(cx);
                        cx.Set(aDate);

                        // ### PARAMETRY IMIR ###
                        XtraReportSerialization.InformacjaZUSIMIR.Params pr = new XtraReportSerialization.InformacjaZUSIMIR.Params(cx);
                        pr.Naliczone = false;
                        cx.Set(pr);

                        var report = new Soneta.Business.UI.ReportResult()
                        {
                            Context = cx,
                            TemplateFileSource = Soneta.Business.UI.AspxSource.Storage,
                            Format = ReportResultFormat.PDF,
                            DataType = typeof(RMUA),
                            TemplateFileName = @"XtraReports/Wzorce użytkownika/imir_podpis.repx"
                        };

                        using (Stream stream2 = rs.GenerateReport(report))
                        {
                            using (var file = System.IO.File.Create(pathFileNamePDF))
                            {
                                CoreTools.StreamCopy(stream2, file);
                                file.Flush();
                            }

                        }

                        // ### NADAWANIE HASŁA ###
                        PdfDocument document = PdfReader.Open(pathFileNamePDF);
                        PdfSecuritySettings securitySettings = document.SecuritySettings;
                        securitySettings.UserPassword = pracownik.PESEL.Substring(6);
                        document.Save(pathFileNamePDF);

                        listaAtt.Add(new System.Net.Mail.Attachment(pathFileNamePDF));

                        potwierdzenie_imir = true;
                    }
                }

                if (potwierdzenie_pit || potwierdzenie_imir)
                {
                    //SendEmail(pracownik.Kontakt.EMAIL, listaAtt);
                    //SendEmailPotwierdzenie(pracownik, listaAtt);
                    wyslano++;

                    //ZaznaczWyslane(potwierdzenie_pit, numer_pit, potwierdzenie_imir, numer_imir);
                }



            }
            string wiadomosc = "";
            if (wyslano > 0)
                wiadomosc += "Wysłano wiadomości e-mail: " + wyslano + Environment.NewLine;
            if (brak_email > 0)
                wiadomosc += "Brak adresów e-mail: " + brak_email + Environment.NewLine;
            if (brak_deklaracji > 0)
                wiadomosc += "Brak deklaracji dla pracowniów: " + brak_deklaracji + Environment.NewLine;
            if (wyslanych_wczesniej > 0)
                wiadomosc += "Deklaracje wysłane wcześniej (pominięto): " + wyslanych_wczesniej + Environment.NewLine;
            if (blad_report_result > 0)
                wiadomosc += "Błąd tworzenia wydruku: " + blad_report_result;


            return new MessageBoxInformation
            {
                Type = MessageBoxInformationType.Information,
                Text = wiadomosc,
                OKHandler = () => null
            };
        }

        public void SendEmail(string to, List<System.Net.Mail.Attachment> attachment)
        {

            var from = new MailAddress("pit_za_22@alexgroup.pl");
            var subject = "Deklaracja podatkowa";
            var body = "Dzień dobry, <br> W załączeniu wysyłamy deklarację PIT 11 oraz informację IMIR za 2022r.<br><br>" +
                "Załączone dokumenty są wydrukiem komputerowym i nie wymagają podpisu wystawiającego/sporządzającego dokument.<br>" +
                " Hasłem dostępu do pliku PDF jest 5 ostatnich cyfr numeru PESEL. <br> <br> " + Context.Login.Operator.FullName + "<br><br>";

            body += "ALEX Sp. z o.o.  NIP: PL 5422865009 <br> ul. Zambrowska 4A, 16-001 Kleosin, Polska <br><br>";
            body += "Treść tej wiadomości zawiera informacje przeznaczone tylko dla adresata. Jeżeli nie jesteście Państwo jej adresatem bądź otrzymaliście ją przez pomyłkę, " +
                "prosimy o powiadomienie o tym nadawcy oraz trwałe jej usunięcie. / This message contains information intended only for the addressee. " +
                "If you are not the addressee or you have received it by mistake, please inform the sender and permanently delete it.";

            var username = "pit_za_22@alexgroup.pl"; // get from Mailtrap
            var password = "v6B}t0S)a6"; // get from Mailtrap

            var host = "smtp.alexgroup.pl";
            var port = 587;

            var client = new SmtpClient(host, port);
            client.Credentials = new NetworkCredential(username, password);
            client.EnableSsl = true;

            var mail = new MailMessage();
            mail.Subject = subject;
            mail.From = from;
            mail.To.Add(to);                    //  <------ PRODUKCJA
            mail.Body = body;
            mail.IsBodyHtml = true;

            foreach (System.Net.Mail.Attachment at in attachment)
                mail.Attachments.Add(at);

            client.Send(mail);

            //mySmtpClient.Send(myMail);

            //      Parametry skrzynki pit_za_22 @alexgroup.pl

            //      Dane podstawowe
            //      Hasło: v6B}t0S)a6
            //      Pojemność: 100 MB
            //      Włączone zabezpieczenie antyspamowe: tak
            //      Automatyczne odpowiedzi: nie

            //      Adresy serwerów pocztowych
            //      Użytkownik: alexgroup.pl @alexgroup.pl
            //      Serwer poczty przychodzącej (POP3, port 110): alexgroup.pl
            //      Serwer poczty wychodzącej(SMTP, port 587): alexgroup.pl
            //      Serwer usługi IMAP(port 143): alexgroup.pl

            //      Adresy serwerów pocztowych z użyciem bezpiecznego połączenia SSL
            //      Użytkownik: pit_za_22 @alexgroup.pl

            //      Serwer poczty przychodzącej z obsługą SSL (POP3, port 995): poczta22389.kei.pl
            //      Serwer poczty wychodzącej z obsługą SSL(SMTP, port 465) : poczta22389.kei.pl
            //      Serwer usługi IMAP z obsługą SSL(port 993) : poczta22389.kei.pl
        }

        public void SendEmailPotwierdzenie(Pracownik pracownik, List<System.Net.Mail.Attachment> attachment)
        {

            var from = new MailAddress("pit_za_22@alexgroup.pl");
            var subject = "Deklaracja podatkowa - " + pracownik.ImięNazwisko;

            var body = "Potwierdzenie wysłania deklaracji podatkowej do pracownika.<br><br>";
            body += "Pracownik: " + pracownik.ImięNazwisko + "<br>";
            body += "Adres email pracownika: " + pracownik.Kontakt.EMAIL + "<br>";
            body += "Godzina wygenerowania deklaracji: " + Date.Now.ToString() + "<br><br>";
            body += "Osoba generująca deklarację:  " + Context.Login.Operator.FullName + "<br><br><br>";

            body += "ALEX Sp. z o.o.  NIP: PL 5422865009 <br> ul. Zambrowska 4A, 16-001 Kleosin, Polska <br><br>";
            body += "Treść tej wiadomości zawiera informacje przeznaczone tylko dla adresata. Jeżeli nie jesteście Państwo jej adresatem bądź otrzymaliście ją przez pomyłkę, " +
                "prosimy o powiadomienie o tym nadawcy oraz trwałe jej usunięcie. / This message contains information intended only for the addressee. " +
                "If you are not the addressee or you have received it by mistake, please inform the sender and permanently delete it.";

            var username = "pit_za_22@alexgroup.pl"; // get from Mailtrap
            var password = "v6B}t0S)a6"; // get from Mailtrap

            var host = "smtp.alexgroup.pl";
            var port = 587;

            var client = new SmtpClient(host, port);
            client.Credentials = new NetworkCredential(username, password);
            client.EnableSsl = true;

            var mail = new MailMessage();
            mail.Subject = subject;
            mail.From = from;
            mail.To.Add("anna.skorupska@autogas-alex.com");
            //mail.To.Add("dklej@fx2.pl");
            mail.Body = body;
            mail.IsBodyHtml = true;

            foreach (System.Net.Mail.Attachment at in attachment)
                mail.Attachments.Add(at);

            client.Send(mail);

        }

        public void ZaznaczWyslane(bool czy_pit, string numer_pit, bool czy_imir, string numer_imir)
        {
            PIT11_29 deklaracja_pit = null;
            RMUA deklaracja_imir = null;

            using (Session session = Context.Login.CreateSession(false, true))
            {
                DeklaracjeModule dm = DeklaracjeModule.GetInstance(session);

                using (ITransaction t = session.Logout(true))
                {
                    // ustawienie cechy deklaracji na wysłane - true
                    if (czy_pit)
                    {
                        deklaracja_pit = (PIT11_29)dm.Deklaracje.NumerWgNumeruDokumentu[numer_pit];
                        deklaracja_pit.Features["Wysłano e-mail"] = true;

                    }
                    if (czy_imir)
                    {
                        deklaracja_imir = (RMUA)dm.Deklaracje.NumerWgNumeruDokumentu[numer_imir];
                        deklaracja_imir.Features["Wysłano e-mail"] = true;
                    }
                    t.Commit();
                }
                session.Save();
            }
        }
    }
}
