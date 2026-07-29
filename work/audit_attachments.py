import json, re
from pathlib import Path
work=Path(r'D:\brushcustom-email-imap\work')
files=sorted(work.glob('emails_*.json'))
rows=[]
for fp in files:
    data=json.loads(fp.read_text(encoding='utf-8'))
    for m in data:
        body=m.get('body','') or ''
        urls=re.findall(r'https?://\S+', body)
        upload_urls=[u.strip('.,;>)"]') for u in urls if 'wp-content/uploads' in u or u.lower().endswith(('.jpg','.jpeg','.png','.pdf','.stp','.step'))]
        atts=m.get('attachments') or []
        if atts or upload_urls:
            rows.append({
              'source':fp.name,
              'date':m.get('date'),
              'from':m.get('from'),
              'subject':m.get('subject'),
              'attachments':atts,
              'upload_urls':upload_urls,
            })
print(json.dumps(rows, ensure_ascii=False, indent=2))
