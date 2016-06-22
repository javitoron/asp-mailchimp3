<!--#include file="lib/setup_MC.asp"-->
<%
    dim list_id
    list_id = "0b43e32f1a"

    dim bat
    set bat = new MCBatch

    call bat.init(mc, null)


    public sub addSubscribe(op, email, name, surname)
        dim merge_vars
        Set merge_vars = New aspJSON

        with merge_vars.data
            .add "email_address", email
            .add "status", "subscribed"
            .add "language", "es_ES"
            .add "merge_fields", merge_vars.Collection()
        
            with merge_vars.data("merge_fields")
                .add "FNAME", name
                .add "LNAME", surname
'                .add "MMERGE3", Request.Form("tratamiento")
 '               .add "MMERGE4", Request.Form("solicitudes")
  '              .add "MMERGE5", Request.Form("alumnoDe")
   '             .add "MMERGE6", Request.Form("titulos")
    '            .add "MMERGE7", Request.Form("bajas")
            end with
        end with

        call bat.post( op, "lists/" & list_id & "/members", merge_vars)
    end sub

    call addSubscribe( "op1", "xxxx@hotmail.com", "hotmail", "guy")
    call addSubscribe( "op2", "xxxx@gmail.com", "gmail", "guy")

    response.write "<h1>Operations</h1>"
    response.write bat.getOperationsString() & "<br /><br />"

    response.write "<h1>Execute response</h1>"
    response.write bat.execute(null) & "<br /><br />"
%>