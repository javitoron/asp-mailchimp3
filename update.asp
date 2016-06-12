<!--#include file="lib/BF_MC_list.asp"-->
<%
    dim list_id
    list_id = "0b43e32f1a"
    dim subscriber_hash, merge_fields
    subscriber_hash = mc.subscriberHash(Request.Form("email"))

    Set merge_fields = New aspJSON
    Set merge_fields.data("merge_fields") = merge_fields.Collection()
    merge_fields.data("merge_fields").add "FNAME", Request.Form("fname")
    merge_fields.data("merge_fields").add "LNAME", Request.Form("lname")
    merge_fields.data("merge_fields").add "MMERGE3", Request.Form("phone")

    call mc.patch("lists/" & list_id & "/members/" & subscriber_hash, merge_fields, null)
    
    if mc.success() then
        response.write "<p>Mailchimp suscriptor updated</p>"
    else
        response.write "<p>Error updating Mailchimp suscriptor. " & _
            mc.getLastHTTPError() & "</p>"

        Response.Write "<p>" & _
            "<strong>" & mc.getAPIErrorTitle() & "</strong>: " & _
            mc.getAPIErrorDetail() & "</p>"
    end if              
%>
