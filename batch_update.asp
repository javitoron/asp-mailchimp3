<!--#include file="lib/setup_MC.asp"-->
<%
    dim list_id
    list_id = "0b43e32f1a"

    dim bat
    set bat = new MCBatch

    call bat.init(mc, null)


    public sub addUpdate(op, email, lname)
        dim subscriber_hash, merge_fields
        subscriber_hash = mc.subscriberHash(email)

        Set merge_fields = New aspJSON
        Set merge_fields.data("merge_fields") = merge_fields.Collection()
        'merge_fields.data("merge_fields").add "FNAME", name
        merge_fields.data("merge_fields").add "LNAME", lname
 
        call bat.patch( op, "lists/" & list_id & "/members/" & subscriber_hash, _
            merge_fields)
    end sub

    call addUpdate( "op1", "xxxx@hotmail.com", "person")
    call addUpdate( "op2", "xxxx@gmail.com", "person")
    response.write "<h1>Operations</h1>"
    response.write bat.getOperationsString() & "<br /><br />"

    response.write "<h1>Execute response</h1>"
    response.write bat.execute(null) & "<br /><br />"
    
%>