# bigcommerce-api-vba
API access to [Bigcommerce](https://bigcommerce.com) store APIs using
Visual Basic for Applications (VBA).
BigCommerce is a full-featured e-commerce platform used by many online stores.
This repository allows access to the API from Microsoft Office macros.
This is useful for generating custom reports from Microsoft Access or
Excel.  This library could be extended to allow updates of the store.

The API for BigCommerce is full-featured and allows the creation of
apps that can update and fetch all important data for a store.  See
the [API documentation](https://developer.bigcommerce.com/api-docs) for full information.

There is already a [BigCommerce GitHub repository](https://github.com/bigcommerce).
It includes API libraries for Python, PHP, and Ruby, but not VBA.

BigCommerce is invited to make VBA API for themselves, and is welcome to
use this code to help them do so.

## Disclaimer!

This repository exists to share work I have done that is useful in my
own store and might be useful to others.  No claim is made to correctness,
and it should be assumed to have undiscovered bugs.  It is not remotely complete,
has weak error checking, has not been thoroughly tested, is poorly documented,
and should not
be used for commercial applications without a thorough code review and
extensive testing.

This repository and all code is completely unsupported.  It may be abandoned,
deleted, or handed over to BigCommerce at any time.

## Current state

This library can connect to and fetch data from the Catalog (Products)
and Orders endpoints, including full filtering for V3 requests.
It automatically handles paging, caching, and enumeration for large array
requests.

It does not (yet) handle rate limiting, concurrency, PUT or POST requests,
most of the API endpoints, or error checking.

It has been tested only for Office 2016 on Windows 10.

## Usage

### Setup

In your VBA project, such as an Excel spreadsheet, include references to
"Microsoft XML, v6.0" and "Microsoft Script Control 1.0".  To do this:

1. Open the "View Code" window
1. Select the "Tools->References..." menu item
1. Check the boxes for the two references
1. Click OK.

Add the JsonConverter module to the project.  To do this,

1. Download the file JsonConverter.bas from the [VBA-tools/VBA-JSON](https://github.com/VBA-tools/VBA-JSON) repository
1. From your VBA project code window, select the menu item "File->Import File..."
1. Select the JsonConverter.bas file
1. Click the Open button

Download the following files from this repository, and add them to your project as above:

* BCFilter.cls
* BCProduct.cls
* BCRequest.cls

Collect credentials for your store.  See [Obtaining Store Credentials](https://developer.bigcommerce.com/api-docs/getting-started/authentication/rest-api-authentication#obtaining-store-api-credentials#obtaining-store-api-credentials).
You will need a client ID, access token, and store hash.

### Functionality

Use a BCRequest object to manage a connection to your store.  Call the BCRequest.Init() method
to provide it with credentials.

Create any objects required as arguments to the API call, such as filters.  This API
uses JSON for payloads in requests and responses.  These are handled via JsonConverter library calls.
This library represents JSON data as Scripting.Dictionary objects, arrays, strings,
Date types, and long integers.

Use the BCRequest object to specify the API call you wish to make (such as MyRequest.GetOrders).
Depending on
the exact endpoint you use, the actual API might not happen until later, on a
cache miss.

Fetch the data.  In the case of API calls that return lists, the results can be enumerated
via the BCRequest.CurDataItem property and BCRequest.NextItem subroutine.

### Sample code

'''VBA
' Make the "Immediate" debugger window visible to view output
Public Sub PrintProducts()
    Dim Request As BCRequest
    Dim Filters As New Collection
    Dim Filter As BCFilter
    
    ' Subsitute your own credentials and store hash
    Const AccessToken = "abcdefghijklmnopqrstuvwxyz01234"
    Const ClientId = "abcdefghijklmnopqrstuvwxyz01234"
    Const StoreHash = "abcdefgh"

    Set Filter = New BCFilter
    Filter.Init "include_fields", BCF_EQ, "sku,id,product_id,is_visible,name"
    Filters.Add Filter
    
    Set Request = New BCRequest
    Request.Init ClientId:=ClientId, AccessToken:=AccessToken, StoreHash:=StoreHash
    Request.GetCatalogProducts limit:=15, Filters:=Filters
    
    Dim Count As Long
    Count = 0
    While Not Request.CurDataItem Is Nothing And Count < 40
        Debug.Print JsonConverter.ConvertToJson(JsonValue:=Request.CurDataItem, Whitespace:=2)
        Request.NextItem
        Count = Count + 1
    Wend
    
End Sub
'''
