# exchange-online-safe-senders
This script provides an easy way for Microsoft partners to add "Safe Sender" lists to their customers' Microsoft Exchange Online tenant via transport rules.

Run the script in the same folder as the CSV file. The CSV file must contain the heading "Domain." Each subsequent row should contain a domain name that will be added to the transport rule as a safe sender (i.e. the Spam Confident Level (SCL) will be set to "-1" only for emails that pass DMARC.

The script prompts for credentials twice: once for Partner Center, and once for Exchange Online. The username from Exchange Online is passed to each Connect-ExchangeOnline cmdlet, so as long as you're authenticating with your partner account credentials, each partner tenant should be updated without having to re-enter credentials again.
