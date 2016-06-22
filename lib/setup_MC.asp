<!--#include file="header_utf8.asp"-->
<!--#include file="MailChimp.asp"-->
<!--#include file="MCBatch.asp"-->
<!--#include file="aspJSON1.17.asp"-->
<%
    dim apikey
    apiKey = "apikey xxxxxxxxxxxxxxx-usx"

    dim mc
    set mc = new MailChimp
    mc.init(apikey)
%>
