```
./Get-MsolUser.py| jq ".[] | select(.UserPrincipalName == \"upndecon@domainedecon.con\")"
```