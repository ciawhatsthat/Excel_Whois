# Excel_Whois
A sloppy whois for a specific csv from excel

Uses MSXML2.XMLHTTP to query https://rdap.arin.net/registry/ip/x.x.x.x and WinHttp.WinHttpRequest.5.1 to query RIPE, LACNIC, APNIC, and LACNIC (redirected from rdap.arin.net)

In this instance, it'll take a csv or emails and IPs and query whois to get Org Name and CIDR of where email came from.

Mostly, just want to keep it here for the whois lookup.
