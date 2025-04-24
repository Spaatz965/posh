param (
    [ValidateRange(1, 36)]
    [int]
    $Month = 3,
    [uri[]]
    $RssFeed = @(
        "https://defenseindepth.libsyn.com/rss",
        "https://davidspark.libsyn.com/cisovendor",
        "https://cisoseries.libsyn.com/rss"
    ),
    [switch]
    $Markdown
)

function ConvertFrom-Html {

    [CmdletBinding()]
    param (
        [ValidateNotNullOrEmpty()]
        [string]
        $HtmlContent,
        [switch]
        $Markdown
    )

    begin {

        $regex = @{

            # [0] Replace non-breaking/unusual space characters
            @'
(?x)

(?:[\u00A0\u1680\u180E\u2000-\u200B\u202F\u205F\u3000\uFEFF]

            # \u00A0: Unicode No-Break Space
            # \u1680: Unicode Ogham Space Mark
            # \u180E: Unicode Mongolian Vowel Separator
            # \u2000-\u200B: Unicode En Quad - Zero Width Space
            # \u202F: Unicode Narrow No-Break Space
            # \u205F: Unicode Medium Mathmatical Space
            # \u3000: Unicode Ideographic Space
            # \uFEFF: Unicode Invalid Character Identifier

|&nbsp;)    # Alternation HTML Character Entitity - Non-Breaking Space
'@ = ' '    # Replace with a single simple space

            # [0.1] replace smart double quotes with straight double quotes
            '[\u201C\u201D\u201E\u201F\u2033\u2036]' = '"'

            # [0.2] replace smart single quotes and appostrophes with straight single quotes
            '[\u2018\u2019\u201A\u201B\u2032\u2035]' = "'"

            # [1] Match/Replace H1 Tags
            '(?:<(h1)(?:\s*[^\r\n>]*\s*)>)(?<text>(?:(?!<\/h1>).)*)(?:<\/h1>)' = "# `${text}"

            # [2] Match/Replace H2 Tags
            '(?:<(h2)(?:\s*[^\r\n>]*\s*)>)(?<text>(?:(?!<\/h2>).)*)(?:<\/h2>)' = "## `${text}"

            # [3] Match/Replace H3 Tags
            '(?:<(h3)(?:\s*[^\r\n>]*\s*)>)(?<text>(?:(?!<\/h3>).)*)(?:<\/h3>)' = "### `${text}"

            # [4] Match/Replace H4 Tags
            '(?:<(h4)(?:\s*[^\r\n>]*\s*)>)(?<text>(?:(?!<\/h4>).)*)(?:<\/h4>)' = "#### `${text}"

            # [5] Match/Replace H5 Tags
            '(?:<(h5)(?:\s*[^\r\n>]*\s*)>)(?<text>(?:(?!<\/h5>).)*)(?:<\/h5>)' = "##### `${text}"

            # [6] Match/Replace H6 Tags
            '(?:<(h6)(?:\s*[^\r\n>]*\s*)>)(?<text>(?:(?!<\/h6>).)*)(?:<\/h6>)' = "###### `${text}"

            # [7] Match/Replace Italics Tags
            '(?:<(i)(?:\s*[^\r\n>]*\s*)>)(?<text>(?:(?!<\/i>).)*)(?:<\/i>)' = "*`${text}*"
            '(?:<(em)(?:\s*[^\r\n>]*\s*)>)(?<text>(?:(?!<\/em>).)*)(?:<\/em>)' = "*`${text}*"
            '(?:<(cite)(?:\s*[^\r\n>]*\s*)>)(?<text>(?:(?!<\/cite>).)*)(?:<\/cite>)' = "*`${text}*"

            # [8] Match/Replace Bold Tags
            '(?:<(b)(?:\s*[^\r\n>]*\s*)>)(?<text>(?:(?!<\/b>).)*)(?:<\/b>)' = "**`${text}**"
            '(?:<(strong)(?:\s*[^\r\n>]*\s*)>)(?<text>(?:(?!<\/strong>).)*)(?:<\/strong>)' = "**`${text}**"
            '(?:<(dfn)(?:\s*[^\r\n>]*\s*)>)(?<text>(?:(?!<\/dfn>).)*)(?:<\/dfn>)' = "**`${text}**"

            # [9] Remove DIV and SPAN Tags
            '(?:<\/?(?:div|span)(?:\s*[^\r\n>]*\s*)>)' = ""

            # [10] Match/Replace Line Breaks
            '(?:[ \t]*<br(?:\s*[^\r\n>]*\s*)(?: *\/)?>\s*)' = "`n"

            # [11] Match/Replace Paragraphs
            '(?:<p(?:\s*[^\r\n>]*\s*)>)(?<text>(?:(?!<\/p>).)*)(?:<\/p>)' = "`${text}`r`n"

            # [12] Remove Images
            '(?:<a[^\r\n>]*>)?(?:<img[^\r\n>]*>)(?:<\/a>)?' = ""

            # [13] Remove Comments
            '(?:<--)(?<text>(?:(?!-->).)*)(?:-->)' = ""

            # [14] Match/Replace Email Links
            '(?:<a(?:\s*[^\r\n>]*\s*)href=\s*")(?:mailto:)(?<email>[^\r\n]+[^\r\n"])(?:">)(?<text>[^\r\n]+[^\r\n<])(?:<\/a>)' =
            """`${text}"" <`${email}>"

            # [15] Remove UTM query parts
            '(?:\?utm[\da-z\/\(\)\!\$\*\+~&;:@=%_-]+)' = ""

            # [16] Remove Unordered List Tags
            '(?:<\/?(?:ul)(?:\s*[^\r\n>]*\s*)>)' = "`r`n"

            # [17] Replace Line Separators
            '[\u2028]' = "`n"

        } # regex

        <#
    TODO:
    - Table
    - Blockquote
    - Figure / Figure Caption
    - Clean up (such as double spaces, etc.)
    - Anchor Tags (Add Support for Files)
#>

        if ( $PSBoundParameters['Markdown'] ) {
            $anchorReplace = "[`${label} (`${domainName}) ](`${url})"
        }
        else {
            $anchorReplace = "**`${label} (`${domainName}) <`${url}>**"
        }

        # Regular expression to extract URLs from anchor tag
        $regex.Add(
@'
(?x)                                             # option - do not capture whitespace or anything following a # sign

(?:<a                                            # START <--ncg--> first part of opening anchor tag
(?:\s*[^\r\n>]+\s*)?                             # <--ncg--> optional space, characters, space to account
                                                    # for attributes
(?:href\s*=\s*")                                 # <--ncg--> literal match 'href="' with optional spaces
)                                                # END <--ncg--> first part of opening anchor tag

(?<url>                                          # START <url> capture group

(?<scheme>\b                                     # START <scheme> capture group
(?:ht|f)                                         # <--ncg--> options to allow for http, https, ftp, and ftps
(?:tps?\b)                                       # <--ncg--> finishing scheme
)                                                # END <scheme> capture group

(?::\/\/)                                        # <--ncg--> literal match '://'

(?<authority>                                    # START <authority> capture group

(?<fqdn>                                         # START <fqdn> capture group

(?<subDomain>                                    # START <subDomain> capture group
(?:\b                                            # start <--ncg-->
(?<!-)                                           # negative look behind - do not start with hyphen '-'
(?:[a-zA-Z\d-]{1,63})                            # one to 63 instances of alpha, numeric, or hyphen
(?!-)                                            # negative look ahead - do not end with hyphen '-'
(?:\.)                                           # literal match for period '.'
){0,10}                                          # zero to ten instances of pattern
)                                                # END <subDomain> capture group

(?<domainName>                                   # START <domainName> capture group

(?:                                              # START <--ncg--> for 2nd level domain options

(?:\b                                            # start of option 1 with a word boundary
(?:[a-zA-Z\d]{2})\b                              # two instances of alpha, numeric
)\.                                              # literal match on period '.'
(?!\b[a-zA-Z\d]{2}                               # two instances of alpha, numeric
(?!\.)\b)|                                       # negative look ahead - do not end with period '.'

(?:\b                                            # start of option 2 with a word boundary
(?:[a-zA-Z\d]{2})                                # two instances of alpha, numeric
(?=\b[a-zA-Z\d]{2}\.[a-zA-Z\d]{2}\b)             # two instances of alpha/numeric with period between
)|

(?:\b                                            # start of option 3 with a word boundary
(?<!-)                                           # negative look behind - do not start with hyphen '-'
(?:[a-zA-Z\d-]{3,63})                            # three to 63 instances of alpha, numeric, or hyphen
(?!-)\b)\.|                                      # negative look ahead - do not end with hyphen '-'

(?:\b                                            # start of option 4 with a word boundary
(?:[a-zA-Z\d]{1})\b)\.                                   # one instance of alpha, numeric, or hyphen

)                                                # END <--ncg--> for 2nd level domain options

(?<eTld>                                         # START <eTld> capture group (effective top level domain) with two options

(?:\b[a-z]{2}\.)?                                # start <--ncg--> and option 1: Optional two letter and period and a two letter country code top level domain
(?<ccTld>\b[a-z]{2}\b                            # <ccTld> capture group (country code top level domain)
(?!\.)                                           # negative look ahead - do not end with period '.'
)|                                               # end Option 1

(?:\b[a-z]{3,30}\b                               # option 2: three to thirty letter top level domain
(?!\.)                                           # negative look ahead - do not end with period '.'
)                                                # end Option 2

)                                                # END <eTld> capture group (effective top level domain) with two options

)                                                # END <domainName> capture group

)                                                # END <fqdn> capture group

(?::                                             # start <--ncg--> optional port number
(?<port>                                         # START <port> capture group
(?:[1-6]\d{4}|\d{1,4})                           # Option of five digits starting with 1-6 or 1-4 digits
)                                                # END <port> capture group
)?                                               # end <--ncg--> optional port number

)                                                # END <authority> capture group

(?<pathName>\/                                   # START <pathName> capture group - optional to capture path part
(?:[\w\.:%~-]+\/)*                               # <--ncg--> zero or more path parts
(?:                                              # start <--ncg--> options for path completion
(?<fileName>[\w\.%~-]+\.[\w\.%~-]+)|             # option 1: <fileName> capture group
(?:[\w:%~-]+)                                    # option 2: path content with no periods
)?                                               # end <--ncg--> options for path completion
)?                                               # END <path> capture group - optional to capture path part

(?:                                              # start <--ncg--> options for query part
(?<query>\?                                      # START <query> capture group (option 1)
(?!utm)                                          # negative look behind - do not start with literal match 'utm'
(?:[\w\.\(\)\*\+\/\$;=:@&%~!'-]+)                # query components
)|                                               # END <query> capture group

(?:\?[\w\.\(\)\*\+\/\$;=:@&%~!'-]+)              # option 2 <--ncg--> capture query allowing for utm strings

)?                                               # end <--ncg--> options for query part

(?<fragment>\#[\w\.\$\(\)\*\+\/\?~!';=:@&%-]+)?  # <fragment> capture group - one or more characters authorized for a fragment

)                                                # END <url> capture group

"(?:\s*[^\r\n>]+\s*)?>                           # completion of opening anchor tag allowing for attributes

(?<label>\s*                                     # START <label> capture group
(?!<img)                                         # negative look behind - do not start with an img tag
(?:                                              # start <--ncg--> allowable text
(?!<\/a>).                                       # negative look behind - allows for any characters not following a closing anchor tag
)*                                               # end start <--ncg--> allowable text
)                                                # END <label> capture group

(?:<\/a>)                                        # <--ncg--> closing anchor tag
'@, $anchorReplace )

        $listItemPattern = '\s*(?:<li(?:\s*[^\r\n>]*\s*)>)(?:\s*<p[^>]*>\s*)?(?<text>(?:(?!<\/li>).)*)(?:<\/p>)?(?:<\/li>)(?:<\/p>)?'

        $orderedListPattern = '(?:<ol(?:\s*[^\r\n>]*\s*)>)(?<text>(?:(?!<\/ol>).)*)<\/ol>'

        $listItemSplit = '(?:</p>)?\s*-\s*'

    } # begin

    process {

        # Regular expression to extract list items from unstructured lists and prefix with a hyphen

        $htmlContent = $htmlContent -replace $listItemPattern, "- `$1`r`n"

        # Regular expression to extract list items from structured lists and prefix with sequential numbers
        $listContent = ( $htmlcontent | select-string -pattern $orderedListPattern -allmatches ).matches

        if ( $null -ne $listContent ) {

            $listItems = $listContent.Groups[1].Value -split $listItemSplit | where-object { $_ -ne "" }

            $numberedItems = for ( $i = 0; $i -lt $listItems.Count; $i++ ) {
                "$( $i + 1 ). $( $listItems[$i] )"
            }

            $listContent = $numberedItems -join " "

            $htmlContent = $htmlContent -replace $orderedListPattern, $listContent

        }

        foreach ( $key in $regex.Keys ) {

            $htmlContent = $htmlContent -replace $key, $regex[$key]

        } # foreach $Key in $regex.Keys

        # Replace HTML entities with their respective characters
        $plainText = [System.Web.HttpUtility]::HtmlDecode($htmlContent)

        # Display plain text
        Write-Output $plainText

    } # process

    end {

    } # end

} # ConvertFrom-Html

function Get-XMRssData {
    <#
.SYNOPSIS
    Pull RSS feed data from a list of RSS feeds and return the data as a PowerShell object.

.DESCRIPTION
    Pull RSS feed data from a list of RSS feeds and return the data as a PowerShell object.
    The function will return the following properties:
        - YearAndMonth
        - episodeDate
        - podcastTitle
        - episodeTitle
        - episodeDuration
        - episodeEntry
        - podcastCopyright
        - podcastSummary
        - podcastAuthor
        - podcastOwner
        - episodeLink
        - episodeDescription
        - episodeKeywords

.PARAMETER  Month
    The number of months from the current date to look back for RSS feed items.
    The default value is 3 months.

.PARAMETER  RssFeed
    An array of URIs for the RSS feeds to pull data from.
    The default values are:
        - https://defenseindepth.libsyn.com/rss
        - https://davidspark.libsyn.com/cisovendor
        - https://cisoseries.libsyn.com/rss

.PARAMETER  Markdown
    A switch parameter that indicates whether to convert HTML content to Markdown format.
    If this parameter is specified, the function will convert HTML content to Markdown format.
    If this parameter is not specified, the function will return the HTML content as-is.

.NOTES
    Additional information about the function or script.
    File Name      : Get-XMRssData.ps1
    Author         : Mark Christman
    Github Repo    : https://github.com/Spaatz965/posh
    Version        : 2.0
    Date           : 24 April 2025

    DISCLAIMER
    The sample scripts are not supported under any support program or service.
    The sample scripts are provided AS IS without warranty of any kind. Author
    further disclaims all implied warranties including, without limitation, any
    implied warranties of merchantability or of fitness for a particular purpose.
    The entire risk arising out of the use or performance of the sample scripts
    and documentation remains with you. In no event shall the authors, or anyone
    else involved in the creation, production, or delivery of the scripts be
    liable for any damages whatsoever (including, without limitation, damages
    for loss of business profits, business interruption, loss of business
    information, or other pecuniary loss) arising out of the use of or inability
    to use the sample scripts or documentation, even if the author has been
    advised of the possibility of such damages.
#>

    [CmdletBinding()]
    param (
        [ValidateRange(1, 36)]
        [int]
        $Month = 3,
        [uri[]]
        $RssFeed = @(
            "https://defenseindepth.libsyn.com/rss",
            "https://davidspark.libsyn.com/cisovendor",
            "https://cisoseries.libsyn.com/rss"
        ),
        [switch]
        $Markdown
    )

    begin {
        Write-Verbose "[$((get-date).TimeOfDay.ToString()) BEGIN   ] "
        Write-Verbose "[$((get-date).TimeOfDay.ToString()) BEGIN   ] Starting: $($MyInvocation.MyCommand)"
        Write-Verbose "[$((get-date).TimeOfDay.ToString()) BEGIN   ] Execution Metadata:"
        # TODO: Add Get-XMNEnvironmentMetaData helper function
        # $EnvironmentData = Get-XMNEnvironmentMetaData
        # Write-Verbose $EnvironmentData

        $startDate = (Get-Date -Day 1 -Hour 0 -Minute 0 -Second 0).AddMonths(-1 * $Month)

    } # begin

    process {
        Write-Verbose "[$((get-date).TimeOfDay.ToString()) PROCESS ] "

        foreach ( $Feed in $RSSFeed ) {

            $FeedData = [xml]( (invoke-webrequest $Feed).Content )

            foreach ( $item in $FeedData.rss.channel.item ) {

                $descriptionParams = @{
                    'htmlContent' = $item.description.InnerText
                }

                if ( $PSBoundParameters['Markdown'] ) {
                    $descriptionParams['Markdown'] = $true
                }

                if ( [datetime]$item.pubDate -ge $startDate ) {

                    if ( $item.title -is [array] ) {
                        $title = $item.title[0]
                    }
                    else {
                        $title = $item.title
                    }

                    $pubDate = Get-Date $item.pubDate -format "dddd, dd MMMM yyyy"
                    $minutes, $seconds = $item.duration -split ':'
                    if ( $minutes -eq "" ) {
                        $minutes = 0
                    }
                    if ( $seconds -eq "" ) {
                        $seconds = 0
                    }
                    $duration = (New-TimeSpan -Minutes $minutes -Seconds $seconds).ToString()

                    $properties = [ordered]@{
                        'YearAndMonth'       = Get-Date $item.pubDate -format "yyyyMM"
                        'episodeDate'        = [datetime]$item.pubDate
                        'podcastTitle'       = $FeedData.rss.channel.title
                        'episodeTitle'       = $title
                        'episodeDuration'    = $duration
                        'episodeEntry'       = "$($FeedData.rss.channel.title) episode: $title, published on: $pubDate - duration: $duration"
                        'podcastCopyright'   = $FeedData.rss.channel.copyright.InnerText
                        'podcastSummary'     = $FeedData.rss.channel.summary.InnerText
                        'podcastAuthor'      = $FeedData.rss.channel.author
                        'podcastOwner'       = "$($FeedData.rss.channel.owner.name.InnerText) <$($FeedData.rss.cannel.owner.email.InnerText)>"
                        'episodeLink'        = $item.link.InnerText

                        # Replaces are regex to strip html because PowerShell 7.x doesn't have an HTML parser
                        'episodeDescription' = convertFrom-html @descriptionParams

                        'episodeKeywords'    = $item.keywords

                    }
                    $Output = New-Object -TypeName PSObject -Property $Properties
                    Write-Output $Output

                }
            }
        }

    } # process

    end {
        Write-Verbose "[$((get-date).TimeOfDay.ToString()) END     ] Ending: $($MyInvocation.Mycommand)"
        Write-Verbose "[$((get-date).TimeOfDay.ToString()) END     ] "

    } # end

} # function Get-XMRssData

#region Function Test Use Cases (Uncomment To Test)

# Call Function with parameters

Get-XMRssData @PSBoundParameters

# Call Function via PipeLine

# Call Function via PipeLine By Name

# Call Function with multiple objects in parameter

# Call Function test Error Logging

#endregion