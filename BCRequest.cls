'' This module sets up a "Connection" with the BigCommerce API and allows
'' processing of the (currently required) endpoints.
''
'' This is single-threaded.  Instructions on faking multi-threading are
'' at http://www.tushar-mehta.com/publish_train/xl_vba_cases/vba_web_pages_services/#_Toc173749360
''
'' This implements paging for requests that allow it.  Pages of results are
'' cached in a dictionary.
''
'' The objects returned are not flat; some of the fields come back as lists
'' or objects.
''
'' Order
'' - billing_address (object)
'' - products (object, but only gives URLs to look up list)
'' - shipping_addresses (object, but only gives URLs to look up list)
''
Option Explicit

'' Error Codes
' Public Const BCRequestErrBase = 845000


'' Info about the assorted endpoints.
''
'' CATALOG
''
'' Catalog Product
'' - categories (array[integer])
'' - related_products (array[integer])
'' - meta_keywords (array-string])
'' - custom_url (object, with fixed fields)
'' - custom_fields (array[object])
'' - bulk_pricing_rules (array[object])
'' - images (array[object])
'' - primary_image (object)
'' - videos (array[object])
'' - options (array[object])
'' - modifiers (array[object])
'' - option_values (array[object])
'' - variants (object)
''
'' ORDERS
''
'' Order
''
'' Ordered Product

' GetCatalogProducts
Private Const cGetCatalogProductsVersion = "v3"
Private Const cGetCatalogProductsPath = "catalog/products"

' GetCatalogVariants
Private Const cGetCatalogVariantsVersion = "v3"
Private Const cGetCatalogVariantsPath = "catalog/variants"

'' ORDERS
'' - billing_address (object)
'' - products (object, but only gives URLs to look up list)
'' - shipping_addresses (object, but only gives URLs to look up list)
Private Const cGetOrdersVersion = "v2"
Private Const cGetOrdersPath = "orders"

'' Order Product
Private Const cGetOrderProductsVersion = "v2"
Private Const cGetOrderProductsPath = "orders/{order_id}/products"


'' Const
Private Const cMaxPageSize As Long = 250
Private Const cMaxCachedPages As Long = 10
Private Const cRESTBase As String = "https://api.bigcommerce.com/stores"

'' Object to manage the HTTP request
Private pRequest As MSXML2.XMLHTTP60

'' These private variables store all the information needed to
'' build the actual API call.
Private pContentType As String
Private pAccept As String
Private pClientId As String
Private pAccessToken As String
Private pEndpointPath As String             ' The endpoint path, such as catalog/products
Private pStoreHash As String
Private pParameters As Scripting.Dictionary
Private pPathParameters As Scripting.Dictionary
Private pHeaders As Scripting.Dictionary
Private pBaseURL As String                  ' URL without the parameters
Private pURL As String                      ' Full URL
Private pAPIVersion As String               ' v2 or v3
Private pFilters As Collection              ' Search filters, to be turned into parameters

'' These private variables store the results of the call
Private pParsedResult As Object
Private pParsedData As Object
Private pParsedMeta As Object
Private pStatus As Long
Private pStatusText As String
Private pErrorText As String

'' These private variables are used to build and manage
'' the cache for paged calls
Private pFiltered As Boolean                ' True for paged GET calls
Private pPageSize As Long                   ' Number of records to fetch
' Private pCurrentPageNum As Long             ' Number of page
Private pPageCacheQueue As Object           ' Holds list of cached pages
Private pMaxCachedPages As Long
Private pCachedPages As Scripting.Dictionary ' Array holding the currently-cached pages
Private pCurrentDataItemNum As Long

'' ---------------------
'' -- Public Properties
'' ---------------------

Public Property Get Status() As Long
    Status = pRequest.Status
End Property

Public Property Get StatusText() As String
    Status = pRequest.StatusText
End Property

Public Property Get ErrorText() As String
    ErrorText = pErrorText
End Property

Public Property Get CurDataItem() As Object
      Set CurDataItem = DataItem(pCurrentDataItemNum)
End Property

Public Property Get DataItem(ItemNumber As Long) As Object
    Dim PageNum As Long
    Dim Page As Object
    Dim ItemNumberInPage
    
    If pCurrentDataItemNum < 1 Then
        Set DataItem = Nothing
    Else
        ' The first page is number 1, not 0
        PageNum = ((ItemNumber - 1) \ pPageSize) + 1
        LoadPage (PageNum)
        If Not pCachedPages.Exists(PageNum) Then
            Set DataItem = Nothing
        Else
            Set Page = pCachedPages.Item(PageNum)
            ' The first item in the page is 1, not 0
            ItemNumberInPage = ((ItemNumber - 1) Mod pPageSize) + 1
            If ItemNumberInPage > Page.Count Then
                Set DataItem = Nothing
            Else
                Set DataItem = pCachedPages.Item(PageNum)(ItemNumberInPage)
            End If
        End If
    End If
End Property





'' ---------------
'' -- Public Subs
'' ---------------

Public Sub NextItem()
    pCurrentDataItemNum = pCurrentDataItemNum + 1
End Sub


Public Sub Reset()
    pCurrentDataItemNum = 1
End Sub



' Intialize the request object.
'
' ClientID, AccessToken, and StoreHash can be strings or ranges.  If
' they are ranges, load the string in from the range value.
' If they aren't specified, use default ranges.
Public Sub Init(ClientId As Variant, _
            AccessToken As Variant, _
            StoreHash As Variant)
    
    ' Set the Client ID
    pClientId = ClientId
    
    ' Set the access token
    pAccessToken = AccessToken
    
    ' Set the store hash
    pStoreHash = StoreHash
    
End Sub

'' GetCatalogProducts()
''
'' Get a list of products
''
'' Note that this doesn't actually make the API call.  The call
'' happens when the the data is requested.
Public Sub GetCatalogProducts(Optional Filters As Collection, _
                                Optional include_fields As Variant, _
                                Optional include As Variant, _
                                Optional limit As Integer = 0, _
                                Optional direction As String = "", _
                                Optional sort As String = "" _
)
    
    Init_Cache
    pAPIVersion = cGetCatalogProductsVersion
    pEndpointPath = cGetCatalogProductsPath
    pFiltered = True
    
    LoadPagedParameters Filters:=Filters, _
                        include_fields:=include_fields, _
                        include:=include, _
                        limit:=limit, _
                        direction:=direction, _
                        sort:=sort
End Sub

'' GetCatalogVariants()
''
'' Get a list of variants (variant products, not VBA variants)
''
'' Note that this doesn't actually make the API call.  The call
'' happens when the the data is requested.
Public Sub GetCatalogVariants(Optional Filters As Collection, _
                                Optional include_fields As Variant, _
                                Optional include As Variant, _
                                Optional limit As Integer = 0, _
                                Optional direction As String = "", _
                                Optional sort As String = "" _
)
    
    Init_Cache
    pAPIVersion = cGetCatalogVariantsVersion
    pEndpointPath = cGetCatalogVariantsPath
    pFiltered = True
    
    LoadPagedParameters Filters:=Filters, _
                        include_fields:=include_fields, _
                        include:=include, _
                        limit:=limit, _
                        direction:=direction, _
                        sort:=sort
End Sub

'' GetOrders()
''
'' Get a list of orders (v2)
''
'' Note that this doesn't actually make the API call.  The call
'' happens when the the data is requested.
Public Sub GetOrders(Optional Filters As Collection, _
                                Optional limit As Integer = 0, _
                                Optional direction As String = "", _
                                Optional sort As String = "" _
)
    
    Init_Cache
    pAPIVersion = cGetOrdersVersion
    pEndpointPath = cGetOrdersPath
    pFiltered = True
    
    LoadPagedParameters Filters:=Filters, _
                        limit:=limit, _
                        direction:=direction, _
                        sort:=sort

End Sub

'' GetOrderProducts()
''
'' Get a list of products from an order (v2)
''
'' Note that this doesn't actually make the API call.  The call
'' happens when the the data is requested.
Public Sub GetOrderProducts(order_id As Variant, _
                                Optional Filters As Collection, _
                                Optional limit As Integer = 0, _
                                Optional direction As String = "", _
                                Optional sort As String = "" _
)
    Init_Cache
    pAPIVersion = cGetOrderProductsVersion
    pEndpointPath = cGetOrderProductsPath
    pFiltered = True
    Set pPathParameters = New Scripting.Dictionary
    pPathParameters.Item("order_id") = order_id
    
    LoadPagedParameters Filters:=Filters, _
                        limit:=limit, _
                        direction:=direction, _
                        sort:=sort

End Sub


'' -----------------------
'' -- Private Methods
'' -----------------------

Private Sub Class_Initialize()
    pContentType = "application/json"
    pAccept = "application/json"
    Set pRequest = New MSXML2.XMLHTTP60
    Set pParameters = New Scripting.Dictionary
    Set pHeaders = New Scripting.Dictionary
    pCurrentDataItemNum = 0 ' Indicates there's no data yet
    pFiltered = False
    Set pFilters = New Collection
End Sub

' Called for endpoints that page the data and thus
' may need a cache.
Private Sub Init_Cache()
    Set pCachedPages = New Scripting.Dictionary
    pMaxCachedPages = cMaxCachedPages
    Set pPageCacheQueue = CreateObject("System.Collections.Queue")
    pPageCacheQueue.Clear
    pCurrentDataItemNum = 1 ' Data starts at 1
End Sub


'' RESTGet()
''
'' Execute a v2 or v3 REST GET on the endpoint specified.
''
'' APIPath      - Path of API endpoint, such as "catalog/products"
'' APIVersion   - API Version, currently "v2" or "v3"
'' StoreHash    - Set the store hash, if present
'' Headers      - Headers, but token and client IDs are put in automatically
'' Parameters   - URL parameters, as name/value dictionary
Private Sub RESTGet(Optional EndpointPath As String = "", Optional APIVersion As String = "", Optional StoreHash As String = "", _
                    Optional Headers As Scripting.Dictionary = Nothing, _
                    Optional Parameters As Scripting.Dictionary = Nothing)
    
    If StoreHash <> "" Then
        pStoreHash = StoreHash
    End If
    
    If EndpointPath <> "" Then
        pEndpointPath = EndpointPath
    End If
    
    If APIVersion <> "" Then
        pAPIVersion = APIVersion
    End If
    
    pBaseURL = cRESTBase & "/" & pStoreHash & "/" & pAPIVersion & "/" & pEndpointPath
    
    ' Build the headers
    If Not Headers Is Nothing Then
        Set pHeaders = Headers
    End If
    If pHeaders Is Nothing Then
        Set pHeaders = New Scripting.Dictionary
    End If
    pHeaders.Item("X-Auth-Client") = pClientId
    pHeaders.Item("X-Auth-Token") = pAccessToken
    pHeaders.Item("Accept") = pAccept
    
     If Not Parameters Is Nothing Then
        Set pParameters = Parameters
    End If
    
    MakeRequest "GET", pBaseURL, pHeaders, pParameters
End Sub

Private Sub RESTPost()
    MakeRequest "POST", URL, Headers, Parameters
End Sub

'' LoadPage()
'' Load a page of data, store it in the cache, and flush
'' any old pages from the cache.
''
'' This assumes that parameters and the URL are already set up,
'' remaining is to call the API endpoint to fetch the data.
Private Sub LoadPage(PageNum As Long)
    ' Debug.Print "Checking page " & PageNum
    
    ' Do nothing if it's already there
    If Not pCachedPages.Exists(PageNum) Then
        ' Debug.Print "Loading page " & PageNum
        ' Fetch the page of data
        pParameters.Item("page") = PageNum
        RESTGet
        If pStatus = 200 Then
            If pAPIVersion = "v2" Then
                Set pParsedData = pParsedResult
                Set pParsedMeta = Nothing
            ElseIf pAPIVersion = "v3" Then
                Set pParsedData = pParsedResult.Item("data")
                Set pParsedMeta = pParsedResult.Item("meta")
            End If
            ' If no results are returned, do nothing
            '' If pParsedMeta.Item("pagination").Item("count") > 0 Then
            If pParsedData.Count > 0 Then
                ' Store it in the cache
                pCachedPages.Add PageNum, pParsedData
    
                ' Mark it in the queue
                pPageCacheQueue.Enqueue (PageNum)
                ' If the cache is maxed, flush stuff
                If pPageCacheQueue.Count > pMaxCachedPages Then
                    pCachedPages.Remove (pPageCacheQueue.Dequeue)
                End If
                ' Debug.Print "Successfully loaded page " & PageNum
                ' Debug.Print JsonConverter.ConvertToJson(JsonValue:=pParsedMeta, Whitespace:=2)
            End If
        Else
            Debug.Print "Page load failed"
        End If
    Else
        ' Debug.Print "Page " & PageNum & "already loaded"
    End If
End Sub



'' Generic base function for making the call.
''
'' This makes the API call, but doesn't return results.  They need to be
'' fetched as a separate step, probably via the xml property.
''
'' Method       - GET or POST
'' URL          - Full URL of the request, not counting the parameters
'' Headers      - Full headers of the request.  If left out, use headers already in pHeaders
'' Parameters   - Full parameters of the request.  If left out, use parameters already in
''              -   pParameters.  Add to URL for GET and body for POST.
Private Sub MakeRequest(Method As String, URL As String, Optional Headers As Scripting.Dictionary, _
                        Optional Parameters As Scripting.Dictionary, _
                        Optional PathParameters As Scripting.Dictionary, _
                        Optional Payload As Variant)
    Const runAsync As Boolean = True
    Const ProcessComplete As Integer = 4
    
    Dim Key As Variant
    Dim Value As String
    Dim ParamStr As String
    Dim FilterStr As String
    Dim Filterv As Variant
    Dim Filter As BCFilter
    Dim i As Long
    
    pBaseURL = URL
    
    ParamStr = ""
    Set pRequest = New MSXML2.XMLHTTP60
    
    If Method <> "GET" And Method <> "POST" Then
        Err.Raise vbObjectError + 374, , "Method must be GET or POST"
    End If
    
    
    If Not Parameters Is Nothing Then
        Set pParameters = Parameters
    End If
    
    If Not pParameters Is Nothing Then
        For Each Key In pParameters.Keys
            If ParamStr <> "" Then
                ParamStr = ParamStr & "&"
            End If
            ParamStr = ParamStr & Key & "=" & pParameters.Item(Key)
        Next Key
    End If
    
    ' Build the filter string
    If Not pFilters Is Nothing Then
        If pFilters.Count > 0 Then
            ' For i = LBound(pFilters) To UBound(pFilters)
            For Each Filter In pFilters
                If ParamStr <> "" Then
                    ParamStr = ParamStr & "&"
                End If
                ' ParamStr = ParamStr & GenFilterString(pFilters(i))
                ParamStr = ParamStr & Filter.Text
            Next
        End If
    End If
    
    ' Modify the path based on the path parameters
    If Not PathParameters Is Nothing Then
        Set pPathParameters = PathParameters
    End If
    If Not pPathParameters Is Nothing Then
        For Each Key In pPathParameters.Keys
            URL = Replace(URL, "{" & Key & "}", pPathParameters.Item(Key))
        Next Key
    End If
    
    ' GET parameters go in the URL
    If Method = "GET" And ParamStr <> "" Then
        pURL = URL & "?" & ParamStr
    Else
        pURL = URL
    End If
    
    pRequest.Open Method, pURL, runAsync

    If Not Headers Is Nothing Then
        Set pHeaders = Headers
    End If
    
    If Not pHeaders Is Nothing Then
        For Each Key In pHeaders.Keys
            pRequest.setRequestHeader Key, pHeaders.Item(Key)
        Next Key
    End If
    
    ' GET calls have no payload
    If Method = "GET" Then
        pRequest.send
    Else
        If ParamStr <> "" Then
            ' POST parameters go in the payload
            pRequest.send ParamStr
        Else
            pRequest.send ' POST, but no parameters
        End If
    End If
    Do While pRequest.readyState <> ProcessComplete
        DoEvents
    Loop
    pStatus = pRequest.Status
    
    ParseJson
    If pRequest.Status <> 200 Then
        pErrorText = pParsedResult(1).Item("message")
        Debug.Print pErrorText
    End If
End Sub

Private Sub ParseJson()

    Set pParsedResult = JsonConverter.ParseJson(pRequest.responseText)
    '' Debug.Print JsonConverter.ConvertToJson(pParsedResult, Whitespace:=2)
End Sub


Private Sub LoadPagedParameters(Optional Filters As Collection, _
                                Optional include_fields As Variant, _
                                Optional include As Variant, _
                                Optional limit As Integer = 0, _
                                Optional direction As String = "", _
                                Optional sort As String = "")
    If Not IsMissing(Filters) Then
        Set pFilters = Filters
    End If
    
    If Not IsMissing(include_fields) Then
        Select Case VBA.VarType(include_fields)
        Case VBA.vbString
            pParameters.Item("include_fields") = include_fields
        Case VBA.Array + VBA.vbString
            pParameters.Item("include_fields") = Join(include_fields, ",")
        Case Else
            Err.Raise 432865, , "include_fields must be a string or array of strings"
        End Select
    End If
    
    If Not IsMissing(include) Then
        Select Case VBA.VarType(include)
        Case VBA.vbString
            pParameters.Item("include") = include
        Case VBA.Array + VBA.vbString
            pParameters.Item("include") = Join(include, ",")
        Case Else
            Err.Raise 432865, , "include must be a string or array of strings"
        End Select
    End If
    
    If limit = 0 Or limit > cMaxPageSize Then
        pPageSize = cMaxPageSize
    Else
        pPageSize = limit
        pParameters.Item("limit") = limit
    End If
    
    If direction <> "" Then
        pParameters.Item("direction") = direction
    End If

    If sort <> "" Then
        pParameters.Item("sort") = sort
    End If
End Sub
