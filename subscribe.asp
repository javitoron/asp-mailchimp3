<!--#include file="lib/setup_MC.asp"-->
<%
    dim list_id
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
        response.write "<p>Mailchimp subscriber added</p>"
    else
        response.write "<p>Error creating Mailchimp subscriber. " & _ 
            mc.getLastHTTPError() & "</p>"

        Response.Write "<p>" & _
            "<strong>" & mc.getAPIErrorTitle() & "</strong>: " & _
            mc.getAPIErrorDetail() & "</p>"

    end if              
%>
