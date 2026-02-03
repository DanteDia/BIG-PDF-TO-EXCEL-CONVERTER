#!/usr/bin/env python
"""Test de conversión Excel a Markdown usando API de Datalab."""

import httpx
import time

API_KEY = 'Z1JLPnKJIAcNosYvpyF-GZrjVizmsf5MsBlNGL7Szuk'
API_URL = 'https://www.datalab.to/api/v1/marker'
file_path = 'TEST_Merge_v8.xlsx'

print('Subiendo Excel a Datalab...')
with open(file_path, 'rb') as f:
    response = httpx.post(
        API_URL,
        files={'file': ('test.xlsx', f, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')},
        data={'output_format': 'markdown', 'mode': 'accurate'},
        headers={'X-API-Key': API_KEY},
        timeout=60.0,
        verify=False
    )

print(f'Status: {response.status_code}')
data = response.json()
print(f'Response: {data}')

if data.get('success'):
    check_url = data['request_check_url']
    print(f'Polling: {check_url}')
    
    for i in range(60):
        time.sleep(3)
        r = httpx.get(check_url, headers={'X-API-Key': API_KEY}, verify=False)
        result = r.json()
        status = result.get('status')
        print(f'Intento {i+1}: {status}')
        
        if status == 'complete':
            md = result.get('markdown', '')
            print(f'\nLongitud: {len(md)} chars')
            print('\n--- Primeros 5000 chars ---')
            print(md[:5000])
            
            # Guardar completo
            with open('excel_to_markdown.md', 'w', encoding='utf-8') as f:
                f.write(md)
            print('\n\nGuardado en excel_to_markdown.md')
            break
        elif status == 'failed':
            print(f'Error: {result.get("error")}')
            break
else:
    print(f'Error inicial: {data}')
