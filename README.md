# barcodeticket
1. My old out of date barcode ticketing system first created in 1999, the system generates a barcode that is linked to a password to confirm ownership. I designed the random number generator instead of using rand(). The reason for this was that rand() uses time as SEED which means it is possible for a same 'random number' to be created when one or more people or user click to purchase at the same exact time. 

2. Actually, I hope others will be able to improve the UX. The problem I am trying to solve is to be able to issue digital or paper printed tickets on demand.

3. The main issue was originally how to I prove I am the owner of this digital ticket or paper ticket ? Say for example, I stole this ticket.

4. The ticket is linked to a password usually ones birthday when one purchase this. So if the ticket is stolen it is unlikely the to pass this verification test.

5. Now the second thing is that if I bought this ticket from a scalper, how do I know this is the only ticket ? I dont because the system is designed to accept only one ticket with the correct pwd. This will discourage people from buying from scalpers as they do not know whether the ticket had not be presented before.

6. Eventually I went on to patent this idea which ended up with US PATENT 7,742,996 but this is confined to Financial Instruments. For example one can print one own cheques and bring it to the store so that upon verification (barcode plus pwd) this financial instrument is spent.Is the same concept except this is now using a financial instrument. 
