<!--#include file="header_utf8.asp"-->
<!--#include file="MailChimp.asp"-->
<!--#include file="aspJSON1.17.asp"-->
<%
    dim apikey
    apiKey = "apikey 07043339c09bfee2e1d7242f8ab29dc4-us7"

    dim mc
    set mc = new MailChimp
    mc.init(apikey)
%>