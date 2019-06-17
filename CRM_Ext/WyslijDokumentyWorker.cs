using System;
using Soneta.Business;
using Soneta.Business.UI;
using Soneta.Handel;
using Soneta.CRM;
using Soneta.Zadania;
using Soneta.Types;
using Soneta.Business.Db;



[assembly: Worker(typeof(CRM_Ext.WyslijDokumentyWorker), typeof(Soneta.Handel.DokHandlowe))]

namespace CRM_Ext
{
    class WyslijDokumentyWorker
    {
        [Context]
        public Context cx { get; set; }
        [Context]
        public Session Session { get; set; }
        [Context]
        public DokumentHandlowy[] dh { get; set; }
      

        //z Contextu wyciągnięcie potrzebnych informacji

        [Action("Wyślij faktury", Priority = 0, Target = ActionTarget.ToolbarWithText)]
        public object StartPDF(Context cx)
        {
            foreach (DokumentHandlowy dokhan in dh)
            {
                return new ReportResult
                {
                    Context = cx,
                    TemplateFileSource = AspxSource.Storage,
                    Format = ReportResultFormat.PDF,

                    TemplateFileName = @"C:\Users\jaroslaw.pitala\Source\Repos\enova\SonetaRaporty\Reports\handel\sprzedaz.aspx",

                    OutputHandler = (stream) =>
                    {
                        var attName = "FV_" + Date.Now.ToString("yyyyMMddHHmmss") + ".pdf";
                        var dirName = "C:\\wydruki\\";
                    // wymagany jest katalog w którym pdf zostanie zapisany, ścieżkę należy dostować do środowiska instalacji
                    var outFileName = System.IO.Path.Combine(dirName, attName);

                        using (var file = System.IO.File.Create(outFileName))
                        {
                            Soneta.Tools.CoreTools.StreamCopy(stream, file);
                            file.Flush();

                        }

                        using (Session s = cx.Login.CreateSession(false, false))
                        {

                            var zm = ZadaniaModule.GetInstance(s);
                            var crm = CRMModule.GetInstance(s);
                            var bm = BusinessModule.GetInstance(s);
                            using (ITransaction transakcja = s.Logout(true))
                            {
                                WiadomoscRobocza wr = new WiadomoscRobocza();
                                crm.WiadomosciEmail.AddRow(wr);
                                wr.Tresc.Add("Witam serdecznie,<br><br>W załączniku przesyłam fakturę numer " + dokhan.Numer + " Prosimy o terminową wpłatę");
                                if (zm.Config.Operatorzy.Aktualny.DomyslneKontoPocztowe != null)
                                {
                                    wr.KontoPocztowe = zm.Config.Operatorzy.Aktualny.DomyslneKontoPocztowe;
                                    wr.Tresc.Add(zm.Config.Operatorzy.Aktualny.Podpis);
                                }
                                Soneta.Business.Db.Attachment zal = new Soneta.Business.Db.Attachment(wr, AttachmentType.Attachments);
                                zal.Name = attName;

                            // stworzenie załącznika z wygenerowanego pliku pdf

                            bm.Attachments.AddRow(zal);
                                System.IO.FileStream strumien = new System.IO.FileStream(outFileName, System.IO.FileMode.Open);
                                zal.LoadFromStream(strumien);
                                zal.LoadIconFromFile(outFileName);

                                transakcja.CommitUI();
                                return null;
                            // return wr;
                        }

                        }


                    }
                };
             
            }
            return null;
        }
    }
}

