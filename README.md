# excelCryptoConnector
An excel addin to load cryptocurrencies data from exchanges or wallets and merge the data into excel queries. 

Right now it only support loading balance (in a query called balance). 

Go to 'cryprto connector' -> 'open account settings' to manage the settings. Enter on this page the api keys from exchanges or public keys from wallets to load the balance automatically. You can enter keys on multiple columns if you have multiple account on one exchange. Use the update button to pull data from the exchanges and wallets (and also update all queries)

It is also possible to manually enter balance in the tab "balance\_manual" to add balances from unsupported exchanges/wallets

sample.xlsx is a sample file that already contains a setup to load the current price from coinmarketcap.com and display the current balance with the current price in bitcoin, usd and euro. It try to match the currencies symbols with coinmarketcap, is the symbol is different between the exchange and coinmarketcap, you can use the tab "manual symbol to coinmarketcap" to enter the relation betwen currency symbol and the currency name on coinmarketcap (it is also possible to enter <USD> or <EUR> to display fiat currencies in the balance table). The tab "accounts balance" show the current value stored on each exchange/wallet.

ETH address if you want to tip me 0xe4A1a6f33dc9C63672eBC6467cBac61D7616a9F5

![screenshot](https://raw.githubusercontent.com/festi/excelCryptoConnector/master/screenshots/balance.png)

![screenshot](https://raw.githubusercontent.com/festi/excelCryptoConnector/master/screenshots/account%20balances.png)
