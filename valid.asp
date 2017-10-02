<%


'==================================
'  INCLUDE FILE FOR FORM VALIDATION
'==================================

'''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  DECLARE GLOBAL VARIABLES
'''''''''''''''''''''''''''''''''''''''''''''''''''''''
DIM errorMSG

validateForm
IF errorMSG <> "" THEN errorForm

'''''''''''''''''''''''''''''''''''''''''''''
'  VALIDATE FORM
'''''''''''''''''''''''''''''''''''''''''''''''''
SUB validateForm
  DIM fieldAttrib, fieldName
     FOR EACH element IN Request.Form
     fieldAttrib   = UCASE( RIGHT( element, 4 ) )
     fieldName = LEFT( element, LEN( element ) - 4 )
	
     IF fieldAttrib = "_REQ"  AND Request.Form( fieldName ) = "" THEN
          errorMSG = errorMSG & " - " & Request.Form( element ) & "<p>"	
     END IF
     IF fieldAttrib = "_VAL"  AND Request.Form( fieldName ) = "" THEN
          errorMSG = errorMSG & " - " & Request.Form( element ) & "<p>"	
     END IF

     IF fieldAttrib = "_VAL" AND Request.Form( fieldName ) <> "" THEN
          SELECT CASE UCASE( Request.Form( element ) )
	CASE "NUMBER"
	     IF NOT isNumeric( Request.Form( fieldName ) ) THEN
		errorMSG = errorMSG & " - " & fieldName & " must be a number.<p>"
	     END IF
	CASE "DATE"
	     IF NOT isDATE( Request.Form( fieldName ) ) THEN
		errorMSG = errorMSG & " - " & fieldName & " must be a date.<p>"
	     END IF
	CASE "CURRENCY"
	     IF NOT isNumeric( Request.Form( fieldName ) ) THEN
		errorMSG = errorMSG & " - " & fieldName & " must be a money amount.<p>"
	     END IF

        CASE "EMAIL"
             IF INSTR(Request.Form(fieldName),"@")=0 OR INSTR(request.Form(Fieldname),".")=0 THEN
                    erorMSG = errorMSG & " - " & fieldName & " must be a valid email address.<p>"
              END IF        


          END SELECT
     END IF
     NEXT
END SUB

'''''''''''''''''''''''''''''''''''''''''''''''''
'  ERROR FORM 
'''''''''''''''''''''''''''''''''''''''''''''''''
SUB errorFORM
	%>
	<html>
	<head><title>Error Form</title></head>
	<body>
	<b>There was a problem with the information you entered:</b>
	<blockquote>
	<%=errorMSG %>
	</blockquote>
	<form method="post" action="<%=Request( "formScript" )%>">
	<% formFields %>
	
	</form>
	
	<%
	Response.End
END SUB

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'  DUMP ALL OF THE FORM FIELDS
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
SUB formFields
     DIM element, fieldAttrib
     FOR EACH element IN Request.Form
          fieldAttrib = UCASE( RIGHT( element, 4 ) )
          IF fieldAttrib <> "_REQ" AND fieldAttrib <> "_VAL" THEN
	%>
	<input name="<%=element%>" 
	type="hidden" 
	value="<%=Server.HTMLEncode( Request.Form( element ) )%>">
	<%
          END IF
     NEXT
END SUB
%>
