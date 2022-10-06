# -*- coding: UTF-8 -*-
from datetime import datetime
import salt.client
import time
import json
import sys
import re
import os

def coleta():
    client = salt.client.LocalClient()
    cmd = """chcp 65001 > NULL 2>&1; \
mkdir C:\\temp\\scripts\\ > NULL 2>&1; \
$Session = New-Object -ComObject "Microsoft.Update.Session"; \
$Searcher = $Session.CreateUpdateSearcher(); \
$historyCount = $Searcher.GetTotalHistoryCount(); \
$Searcher.QueryHistory(0, $historyCount) | Select-Object Title, Date, @{name="Operation"; \
expression={switch($_.operation){ 1 {"Installation"}; 2 {"Uninstallation"}; 3 {"Other"}}}} \
| ConvertTo-Json -Compress \
| Out-File C:\\temp\scripts\\getkb.json -Encoding utf8 \
| Set-Content -Encoding UTF8 C:\\temp\\scripts\\getkb.json; \
Get-Content C:\\temp\\scripts\\getkb.json"""

    # call salt
    jid = client.cmd_async('kernel:Windows', 'cmd.run', [cmd], kwarg={'shell' : 'powershell'}, tgt_type='grain')
    time.sleep(10)
    res = client.get_cache_returns(jid)
    # handling the data
    newkb = {}
    ret_data = {}
    for minion_id in res:
        print(minion_id, '\n')
        minion_patches = []
        patches = res[minion_id]['ret']
        patches = json.loads(patches)
        newpatches = {}
        minion_patches2 = {}
        minion_patches3 = []
        for path in patches:
            # get Time
            date = path.get('Date')
            GETDATE = re.compile('[0-9]{1,}', re.IGNORECASE)
            sdate = GETDATE.search(date)
            sdate = sdate.group(0) if sdate else 'NA'
            date = datetime.fromtimestamp(int(sdate)//1000)
            # get KB
            GETKB = re.compile('KB[0-9]{6,7}', re.IGNORECASE)
            title = path['Title']
            kbl = GETKB.search(title)
            kbl = kbl.group(0) if kbl else 'NA'
            # new obj
            newpatches = {'kb': kbl, 'installed_on': str(date)}

            if not kbl in minion_patches2:
                minion_patches2[kbl] = []

            if kbl in minion_patches2:
                minion_patches2[kbl].append(newpatches)

        for key, value in minion_patches2.items():
            def key_func(k):
                return k['installed_on']

            minion_patches_sorted = sorted(value, key=key_func, reverse=True)
            minion_patches3.append(minion_patches_sorted[0])
    
    newkb[minion_id] = minion_patches3
    return json.dumps(newkb, indent=2)

# Call the script
print(coleta())
