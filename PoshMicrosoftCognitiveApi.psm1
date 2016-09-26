<#
    .SYNOPSIS
    Uses Microsoft cognitive API to spellcheck a sentence
    
    .DESCRIPTION
    Submit a string and run it against Microsofts API to check spelling, grammer
    common errors, capitalization etc
    More here: https://www.microsoft.com/cognitive-services/en-us/bing-spell-check-api
    
    .PARAMETER ApiKey
    Mandatory
    The API key, get yours from below
    https://www.microsoft.com/cognitive-services/en-us/bing-spell-check-api
  
    .PARAMETER Sentence
    Mandatory
    The sentence (or word) to check
      
    .PARAMETER Mode
    Mandatory
    Spell or Proof
    Spell is for a single word, is more aggressive for better search results
    Proof is less aggressive and adds capitalization, basic punctuation, and other features 
    to aid document creation.

    .PARAMETER CorrectMe
    Optional
    Specify to just return corrected results (assumes first result)
    Don't specify to return the whole object

    .PARAMETER HighlightErrors
    Optional
    Specify to write to screen the potential errors
    
    .EXAMPLE
    $Sentence = "I like chese on tosat. It is delcious, espeshially with tomattos and ham!"
    Get-SpellCheck -ApiKey "56e73033cf5ff80c2008c679" -Sentence $Sentence -Mode "proof" -CorrectMe -HighlightErrors
    
    .EXAMPLE
    $Sentence = "pinapple"
    Get-SpellCheck -ApiKey "56e73033cf5ff80c2008c679" -Sentence $Sentence -Mode "Spell" -CorrectMe

    .OUTPUT
    The corrected sentence/word (with -CorrectMe)
    or the spelling object
    
    .NOTES
    Written by Martin Howlett @WillCode4Pizza
    API doc: https://dev.cognitive.microsoft.com/docs/services/56e73033cf5ff80c2008c679/operations/56e73036cf5ff81048ee6727
#>
Function Get-SpellCheck
{
    Param(
        [Parameter(Mandatory=$true)][string]$ApiKey,
        [Parameter(Mandatory=$true)][string]$Sentence,
        [ValidateSet("spell","proof")]
        [Parameter(Mandatory=$true)][string]$Mode,
        [Parameter()][switch]$CorrectMe,
        [Parameter()][switch]$HighlightErrors
    )

    #Build the URL with the parms
    $Uri = "https://api.cognitive.microsoft.com/bing/v5.0/spellcheck/?text=$Sentence&mode=$Mode"
    
    #Build our params for invoke-restmethod
    #Ocp-Apim-Subscription-Key is just your ApiKey
    $Params = @{
        Headers = @{"Ocp-Apim-Subscription-Key" = "$ApiKey"}
        Uri = $Uri
        Method = "Get"
        ContentType = "application/json"
    }

    #Run the query
    $Response = $null
    Try
    {
        $Response = Invoke-RestMethod @Params
    }
    catch
    {
        Write-Output $False
        Write-host "ERROR :: $($_.Exception.Message)"
    }

    #use the suggestions for each response to generate the new sentence
    $NewSentence = $Sentence
    Foreach($Token in $Response.flaggedTokens)
    {
        $Suggestion = $($Token.suggestions[0].suggestion)
        $Word = $($Token.token)
        $NewSentence = $NewSentence -Replace $Word,$Suggestion

        #write to screen if selected
        if($HighlightErrors)
        {
            Write-Host "$($Token.token)" -ForegroundColor Yellow -NoNewline
            Write-Host " could be spelt as " -NoNewline
            Write-Host "$Suggestion" -ForegroundColor Yellow
        }
    }

    #respond with the corrected sentence or the complete response
    if($CorrectMe)
    {
        Write-Output $NewSentence 
    }
    else
    {
        Write-Output $Response 
    }
}

<#
    .SYNOPSIS
    Uses Microsoft API to perform a bing search
    
    .DESCRIPTION
    Runs a Bing search against the Microsofts API. Retrieve web documents indexed by Bing 
    and narrow down the results by result type, freshness and more.
    
    .PARAMETER ApiKey
    Mandatory
    The API key, get yours from below
    https://www.microsoft.com/cognitive-services/en-us/bing-web-search-api
  
    .PARAMETER Query
    Mandatory
    The search query. Supports normal advanced Bing syntax
      
    .PARAMETER Site
    Optional
    Filter your query to a website 
    
    .PARAMETER Count
    Optional
    The number of results to return
    Default is 10

    .PARAMETER Offset
    Optional
    Where to start returning results from
    Default is 0
    
    .PARAMETER Market
    Optional
    Default is en-gb
    The region to search in
    
    .PARAMETER safeSearch
    Optional
    Default is moderate
     
    .PARAMETER OpenInFirefox
    Optional
    If selected, returns the first URL and opens it in Firefox
     
    .EXAMPLE
    Normal search
    Get-BingWebSearch -ApiKey "56e73033cf5ff80c2008c679" -Query "Powershell module"

    .EXAMPLE
    Normal search, open first result in firefox
    Get-BingWebSearch -ApiKey "56e73033cf5ff80c2008c679" -Query "Powershell module" -OpenInFirefox

    .EXAMPLE
    Normal search, open first result in firefox, filter by site stackoverflow
    Get-BingWebSearch -ApiKey "56e73033cf5ff80c2008c679" -Query "Powershell module" -site "stackoverflow.com" -OpenInFirefox
    
    .OUTPUT
    Search object
        
    .NOTES
    Written by Martin Howlett @WillCode4Pizza
    API doc: https://dev.cognitive.microsoft.com/docs/services/56b43eeccf5ff8098cef3807/operations/56b4447dcf5ff8098cef380d
#>
Function Get-BingWebSearch
{
    Param(
        [Parameter(Mandatory=$true)][string]$ApiKey,
        [Parameter(Mandatory=$true)][String]$Query,
        [Parameter()][String]$Site,
        [Parameter()][int]$Count = 10,
        [Parameter()][int]$Offset = 0,
        [Parameter()][String]$Market = "en-gb",
        [Parameter()][String]$safeSearch = "Moderate",
        [Parameter()][switch]$OpenInFirefox
    )

    #if the user has selected a site to filter the query to, add the syntax
    if($site)
    {
        $Query = "$Query site:$Site"
    }

    #Build the URL with the params
    $uri = "https://api.cognitive.microsoft.com/bing/v5.0/search?q=$Query&count=$Count&offset=$Offset&mkt=$Market&safesearch=$safeSearch"
    $Params = @{
        Headers = @{"Ocp-Apim-Subscription-Key" = "$ApiKey"}
        Uri = $uri
        Method = "Get"
        ContentType = "application/json"
    }
    
    $Response = $null
    Try
    {
        $Response = Invoke-RestMethod @Params
    }
    catch
    {
        Write-Output $False
        Write-Verbose "ERROR :: $($_.Exception.Message)"
    }

    #if user selected, open first URL in firefox    
    if($OpenInFirefox)
    {
        $URL = $response.webPages.value[0].url
        & "C:\Program Files (x86)\Mozilla Firefox\firefox.exe" $URL
    }
    
    #return the response
    Write-Output $response
}

<#
    .SYNOPSIS
    Uses Get-BingWebSearch to search hey scripting guy blog
    
    .DESCRIPTION
    Wrapper around Get-BingWebSearch to filter web search to hey scripting guy blog
    will open first result in Firefox
    
    .PARAMETER ApiKey
    Mandatory
    The API key, get yours from below
    https://www.microsoft.com/cognitive-services/en-us/bing-web-search-api
  
    .PARAMETER Query
    Mandatory
    The search query for the hey scripting guy blog
            
    .EXAMPLE
    Get-ScriptingGuyQuery -ApiKey "56e73033cf5ff80c2008c679" -Query "Get-Childitem"

    .EXAMPLE
    Get-ScriptingGuyQuery -ApiKey "56e73033cf5ff80c2008c679" -Query "PSObject"

    .OUTPUT
    Opens first URL found in Firefox
        
    .NOTES
    Written by Martin Howlett @WillCode4Pizza
#>
Function Get-ScriptingGuyQuery
{
    Param(
        [Parameter(Mandatory=$true)][string]$ApiKey,
        [Parameter(Mandatory=$true)][String]$Query
    )

    $Response = Get-BingWebSearch -ApiKey $ApiKey -Site "https://blogs.technet.microsoft.com/heyscriptingguy" -Query $Query -OpenInFirefox
}
