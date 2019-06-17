using Soneta.Business;
using Soneta.Business.UI;
using Soneta.CRM;
using Soneta.Types;
using Soneta.CRM.Config;

[assembly: Worker(typeof(CRM_Ext.MailingZbiorczyWorker), typeof(Soneta.CRM.WiadomosciEmail))]

namespace CRM_Ext
{

    public class MyParameters : ContextBase
    {
      

        public MyParameters(Context ctx) : base(ctx) { }

        [Caption("Konto pocztowe"), Soneta.Tools.Priority(11)]
        public Soneta.CRM.Config.KontoPocztowe kp
        {
            get;
            set;
        }

    }

    
    class MailingZbiorczyWorker
    {
        [Context]
        public Session s { get; set; }

        [Context]
        public WiadomoscEmail[] wr { get; set; }
        [Context]
        public MyParameters Params { get; set; }

        [Action("Wyślij zbiorczo", Mode = ActionMode.SingleSession | ActionMode.Progress, Target = ActionTarget.Toolbar | ActionTarget.Menu | ActionTarget.LocalMenu | ActionTarget.Divider)]
        public void WyslijMaile(Context cx)
        {

            foreach (WiadomoscEmail wiadomosc in wr)
            {
                if (wiadomosc.TypWiadomosci == TypWiadomości.Robocza)
                {
                    
                        using (ITransaction myTransaction = s.Logout(true))
                        {
                           
                            wiadomosc.KontoPocztowe = Params.kp;
                            wiadomosc.Od = Params.kp.Nazwa;

                            MailHelper.SendMessage(wiadomosc);
                            wiadomosc.TypWiadomosci = TypWiadomości.Wysłana;
                            myTransaction.CommitUI();
                        }
                    
                }

            }
        }

    }
}




