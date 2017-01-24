<!--#include file="md5.asp"-->
<%
 ' Super-simple, minimum abstraction MailChimp API v3 wrapper
 ' MailChimp API v3: http://developer.mailchimp.com
 ' Based on PHP version: https://github.com/drewm/mailchimp-api
 '
 ' @author Javier Torón <javitoron@gmail.com>
 ' @version 0.1

class MailChimp

    private api_key
    private api_endpoint

    public verify_ssl
    private request_successful
    private last_http_error
    private api_error_title
    private api_error_detail
    private last_response_headers
    private last_response_body
    private last_request
    private last_posted_fields

    Private Sub Class_Initialize()
        api_endpoint = "https://<dc>.api.mailchimp.com/3.0"
        verify_ssl = true
        request_successful = false
        last_http_error = ""
    End Sub

    Private Sub Class_Terminate()
    End Sub

    ' Create a new instance
    ' @param string api_key Your MailChimp API key
    public sub init(p_api_key)
        api_key = p_api_key

        if p_api_key = "" then
            Response.write "Empty MailChimp API key."
            response.end
        end if

        dim dashRevPos
        dashRevPos = instrrev(api_key, "-us")

        if dashRevPos = 0 then
            Response.write "Invalid MailChimp API key supplied."
            response.end
        end if
        dashRevPos = len(api_key) - dashRevPos

        dim data_center
        data_center = right(api_key, dashRevPos )
        api_endpoint  = replace(api_endpoint, "<dc>", data_center)
    end sub


    'Create a new instance of a Batch request. Optionally with the ID of an existing batch.
    '@param string batch_id Optional ID of an existing batch, if you need to check its status for example.
    '@return Batch            New Batch object.
'    public function new_batch(batch_id)
'        new_batch = new Batch(Me, batch_id)
'    end

     ' Convert an email address into a "subscriber hash" for identifying the 
        '   subscriber in a method URL
     ' @param   string email The subscriber"s email address
     ' @return  string          Hashed version of the input
    public function subscriberHash(email)
        subscriberHash = MD5(lcase(email))
    end function

    'Was the last request successful?
    '@return bool  True for success, false for failure
    public function success()
        success = request_successful
    end function

    'Get the last error returned by the MSXML2.ServerXMLHTTP object
    public function getLastHTTPError()
        getLastHTTPError = last_http_error
    end function

    'Get the title of the last error returned by the API.
    public function getAPIErrorTitle()
        getAPIErrorTitle = api_error_title
    end function

    'Get the details of the last error returned by the API.
    public function getAPIErrorDetail()
        getAPIErrorDetail = api_error_detail
    end function


    'Get HTTP headers of the API response.
    '@return string
    public function getLastResponseHeaders()
        getLastResponseHeaders = last_response_headers
    end function

    'Get HTTP body of the API response.
    '@return JSON string
    public function getLastResponseBody()
        getLastResponseBody = last_response_body
    end function

    'Get an string containing the HTTP headers of the API request.
    '@return string
    public function getLastRequest()
        getLastRequest = last_request
    end function

    'Get an string containing the posted fields of the API request.
    '@return JSON string
    public function getPostedFields()
        getPostedFields = last_posted_fields
    end function

    'Make an HTTP DELETE request - for deleting data
    '@param   string method URL of the API request method
    '@param   JSON string arguments (if any)
    '@param   int timeout Timeout limit for request in seconds
    public sub delete(p_method, p_args, p_timeout)
        if IsNull(p_timeout) then
            p_timeout = 10
        end if

        call makeRequest("delete", p_method, p_args, p_timeout)
    end sub

    'Make an HTTP GET request - for retrieving data
    '@param   string method URL of the API request method
    '@param   JSON string arguments (usually your data)
    '@param   int timeout Timeout limit for request in seconds
    public sub getRequest(p_method, p_args, p_timeout)
        if IsNull(p_timeout) then
            p_timeout = 10
        end if

        call makeRequest("get", p_method, p_args, p_timeout)
    end sub

    'Make an HTTP PATCH request - for performing partial updates
    '@param   string method URL of the API request method
    '@param   JSON string arguments (usually your data)
    '@param   int timeout Timeout limit for request in seconds
    public sub patch(p_method, p_args, p_timeout)
        if IsNull(p_timeout) then
            p_timeout = 10
        end if

        call makeRequest("patch", p_method, p_args, p_timeout)
    end sub
    
    'Make an HTTP POST request - for creating and updating items
    '@param   string method URL of the API request method
    '@param   JSON string arguments (usually your data)
    '@param   int timeout Timeout limit for request in seconds
    public sub post(p_method, p_args, p_timeout)
        if IsNull(p_timeout) then
            p_timeout = 10
        end if

        call makeRequest("post", p_method, p_args, p_timeout)
    end sub
    
    'Make an HTTP PUT request - for creating new items
    '@param   string method URL of the API request method
    '@param   JSON string arguments (usually your data)
    '@param   int timeout Timeout limit for request in seconds
    public sub putRequest(p_method, p_args, p_timeout)
        if IsNull(p_timeout) then
            p_timeout = 10
        end if

        call makeRequest("put", p_method, p_args, p_timeout)
    end sub
    
    'Performs the underlying HTTP request. Not very exciting.
    '@param  string http_verb The HTTP verb to use: get, post, put, patch, delete
    '@param  string method The API method to be called
    '@param   JSON string arguments (parameters to be passed)
    '@param int timeout
    private sub makeRequest(http_verb, method, args, timeout)
        if IsNull(timeout) then
            timeout = 10
        end if

        dim url, req, params
        url = api_endpoint & "/" & method
        last_http_error = ""
        request_successful = false
        dim post_fields
        post_fields = ""

        last_request = "method: " & http_verb & ", path: " & method & _
            ", url: " & url & ", timeout: " & timeout

        set req = Server.CreateObject("MSXML2.ServerXMLHTTP.6.0")
        'Option SXH_OPTION_URL_CODEPAGE = 0
        req.setOption 0, 65001
        req.setTimeouts timeout * 500, timeout * 500, timeout * 1000, _
            timeout * 1000 'ms - resolve, connect, send, receive
        req.open http_verb, url & params, false ', "username", "password"
        req.setRequestHeader "Authorization", api_key
        req.setRequestHeader "Content-Type", "application/json"

        if http_verb = "post" or http_verb = "patch" or http_verb = "put" then
            post_fields = args.JSONoutput()
        end if
            
        last_posted_fields = post_fields
        req.send post_fields

        last_response_headers = req.getAllResponseHeaders()
        last_response_body = req.responseText

        select case http_verb
            case "get", "post", "put", "delete", "patch"
                if req.status >= 400 and req.status <= 599 then
                    request_successful = false
                    last_http_error = _
                        "Error: " & req.Status & " - " & req.statusText

                    dim myJSON
                    set myJSON = new aspJSON
                    myJSON.loadJSON_from_string(last_response_body)

                    api_error_title = myJSON.data.item("title")
                    api_error_detail = myJSON.data.item("detail")
                else
                    request_successful = true
                    last_http_error = ""
                end if

        end select
    end sub

 end class
%>
