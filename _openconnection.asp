
<%
   set conn = Server.CreateObject("adodb.connection")

   ' DSNless connection to Access Database
   'DSN = "DRIVER={Microsoft Access Driver (*.mdb)}; PWD=;THREADS=25; MAXBUFFERSIZE=2048;DBQ=" & server.mappath("dbs\cc.mdb")
   'conn.Open DSN

   'with jet ole db


   DSN = "Provider=Microsoft.Jet.OLEDB.4.0;persist security info=false;Data Source=" & server.mappath("dbs3789729794\cc.mdb") 
   'DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & server.mappath("dbs\cc.mdb")
   'DSN = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=c:\inetpub\wwwroot\demo\dbs\cc.mdb;"
   conn.Open DSN
   
   ' Connection with SQL Server
  'DSN = "driver={sql server};server=X6c1x3;uid=sa;pwd=;database=test"
   'DSN = "driver={sql server};server=athena;uid=internet;pwd=internet;database=testing"
   'conn.open DSN



'conn.Open "Provider=SQLOLEDB; Data Source = (local); Initial Catalog = test; User Id = sa; Password=" 
  
  
   'DSN= "Provider=MSDASQL;" & _  
     '              "Driver={SQL Server};" & _
     '              "Server=myServerName;" & _
     '              "Database=myDatabaseName;" & _
     '              "Uid=myUsername;" & _
     '              "Pwd=myPassword;"
     'Provider=sqloledb

    'DSN= "Provider=sqloledb;" & _ 
     '              "Data Source=(local);" & _
     '              "Initial Catalog=test;" & _
     '              "User Id=sa;" & _
      '             "Password=;"

    ' conn.open DSN

' Because of the nature of MSSQL and Jet does not work, one has to revert back to driver method
' hence one need to change the date format as below. {ts is not for JET. For JET use "#". JET is
' usually use for ACCESS database. If you are using MSSQL use {ts

Function PrepDate(DateToConvert)
	y = year(DateToConvert)
	m = month(DateToConvert)
	if len(m) = 1 then m = "0" & m
	d = day(DateToConvert)
	if len(d) = 1 then d = "0" & d
        'strDate = "{ts '" & y & "-" & m & "-" & d & " " & FormatDateTime(DateToConvert, 4) & ":00'}"
        'strDate = "" & y & "-" & m & "-" & d & " " & FormatDateTime(DateToConvert, 4) & ":00 "
        strDate = "#" & y & "-" & m & "-" & d & " " & FormatDateTime(DateToConvert, 4) & ":00 # "
	PrepDate = strDate
end function

%>
