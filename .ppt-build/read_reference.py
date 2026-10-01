import re,html,json,sys
from pathlib import Path
t=Path('D:/hems-search/output/Energy_Pack_Gateway_Technology_2027-2030.html').read_text(encoding='utf-8')
sections={}
for ident,body in re.findall(r'<section id="([^"]+)">(.*?)</section>',t,re.S):
 body=re.sub(r'<(?:br|/p|/li|/tr|/td|/dd|/dt|/h[1-6]|/div)[^>]*>', '\n',body)
 sections[ident]=html.unescape(re.sub(r'<[^>]+>',' ',body))
for i in sys.argv[1:]: print(i+'\n'+sections[i]+'\n')
Path('D:/hems-search/.ppt-build/reference-sections.json').write_text(json.dumps(sections,ensure_ascii=False,indent=2),encoding='utf-8')
