using System.Text.RegularExpressions;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using Microsoft.Exchange.Data.Transport;
using Microsoft.Exchange.Data.Transport.Smtp;

namespace Microsoft.Exchange.DynamicAlias{
    [RunInstaller(true)]
    public class DynamicAliasEventLogInstaller : Installer{
        private EventLogInstaller AliasEventLogInstaller;

        public DynamicAliasEventLogInstaller(){
            AliasEventLogInstaller = new EventLogInstaller();
            AliasEventLogInstaller.Source = "Microsoft.Exchange.DynamicAlias";
            AliasEventLogInstaller.Log = "Application";
            Installers.Add(AliasEventLogInstaller);
        }
    }

    public sealed class DynamicAliasFactory : SmtpReceiveAgentFactory {
        public override SmtpReceiveAgent CreateAgent(SmtpServer server){
            return new DynamicAliasAgent(server);
        }
    }

    public class DynamicAliasAgent : SmtpReceiveAgent{
        private readonly string EventLogClass = "Microsoft.Exchange.DynamicAlias";
        private readonly SmtpServer Server;
        public DynamicAliasAgent(SmtpServer ExchangeServer){
            Server = ExchangeServer;
            OnRcptCommand += RcptToHandler;
        }

        public void RcptToHandler(ReceiveCommandEventSource source, RcptCommandEventArgs rcptArgs){
            string RecipientAddress = rcptArgs.RecipientAddress.ToString().ToLower();
            RoutingAddress CleanRecipientAddress = (RoutingAddress)new Regex("\\+\\S+\\@").Replace(RecipientAddress, "@");
            if(null != Server.AddressBook.Find(CleanRecipientAddress)){
                EventLog.WriteEntry(EventLogClass, "Alias valid, mapping " + RecipientAddress + " >> " + CleanRecipientAddress.ToString(), EventLogEntryType.Information, 1);
                rcptArgs.RecipientAddress = CleanRecipientAddress;
                return;
            }

            return;
        }
    }
}
