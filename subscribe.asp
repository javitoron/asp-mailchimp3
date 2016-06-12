<!--#include file="lib/BF_MC_list.asp"-->
<%
    dim list_id, apikey
    list_id = "0b43e32f1a"
    dim subscriber_hash, merge_fields, merge_vars
    subscriber_hash = mc.subscriberHash(Request.Form("email"))

    Set merge_vars = New aspJSON

    with merge_vars.data
        .add "email_address", Request.Form("email")
        .add "status", "subscribed"
        .add "language", "es_ES"
        .add "merge_fields", merge_vars.Collection()
        
        with merge_vars.data("merge_fields")
            .add "FNAME", Request.Form("fname")
            .add "LNAME", Request.Form("lname")
            .add "MMERGE3", Request.Form("phone")
        end with
    end with
 
    call mc.post("lists/" & list_id & "/members", merge_vars, null)
    
    if mc.success() then
        response.write "<p>Suscriptor Mailchimp creado</p>"
    else
        response.write "<p>Error al crear el suscriptor en Mailchimp. " & _ 
            mc.getLastError() & "</p>"

        dim myJSON
        set myJSON = new aspJSON
        myJSON.loadJSON_from_string(mc.getLastResponseBody())

        Response.Write "<p>" & _
            "<strong>" & myJSON.data.item("title") & "</strong>: " & _
            myJSON.data.item("detail") & "</p>"

    end if              
%>
