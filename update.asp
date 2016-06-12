<!--#include file="lib/BF_MC_list.asp"-->
<%
    dim list_id, apikey
    list_id = "0b43e32f1a"
    dim subscriber_hash, merge_fields
    subscriber_hash = mc.subscriberHash(Request.Form("email"))

    Set merge_fields = New aspJSON
    Set merge_fields.data("merge_fields") = merge_fields.Collection()
    merge_fields.data("merge_fields").add "FNAME", Request.Form("fname")
    merge_fields.data("merge_fields").add "LNAME", Request.Form("lname")
    merge_fields.data("merge_fields").add "MMERGE3", Request.Form("tratamiento")
    merge_fields.data("merge_fields").add "MMERGE4", Request.Form("solicitudes")
    merge_fields.data("merge_fields").add "MMERGE5", Request.Form("alumnoDe")
    merge_fields.data("merge_fields").add "MMERGE6", Request.Form("titulos")
    merge_fields.data("merge_fields").add "MMERGE7", Request.Form("bajas")
 
    call mc.patch("lists/" & list_id & "/members/" & subscriber_hash, merge_fields, null)
    
    if mc.success() then
        response.write "<p>Suscriptor Mailchimp modificado</p>"
    else
        response.write "<p>Error al modificar el suscriptor en Mailchimp</p>"
     '   echo "<p>" . $MailChimp->getLastError() . "</p>";
    end if              
%>