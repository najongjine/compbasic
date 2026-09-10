from pathlib import Path
from zipfile import ZipFile, ZIP_DEFLATED
from lxml import etree as E
from copy import deepcopy
p=Path('D:/compbasic/한글기초/예제/토끼고기_간장조림_만들기.hwpx')
with ZipFile(p) as z: data={n:z.read(n) for n in z.namelist()}
ns={'hp':'http://www.hancom.co.kr/hwpml/2011/paragraph','hh':'http://www.hancom.co.kr/hwpml/2011/head','hc':'http://www.hancom.co.kr/hwpml/2011/core'}
h=E.fromstring(data['Contents/header.xml']); s=E.fromstring(data['Contents/section0.xml'])
tbl=s.find('.//hp:tbl',ns); widths=[12472,10488,25228]
tbl.find('hp:sz',ns).set('width','48188')
tbl.find('hp:sz',ns).set('height','13600')
for row in tbl.findall('hp:tr',ns):
 for i,cell in enumerate(row.findall('hp:tc',ns)):
  sz=cell.find('hp:cellSz',ns);sz.set('width',str(widths[i]));sz.set('height','1700')
  cell.set('hasMargin','1')
  mg=cell.find('hp:cellMargin',ns)
  for k,v in {'left':'400','right':'400','top':'250','bottom':'250'}.items():mg.set(k,v)
for x in h.findall('.//hh:borderFill',ns):
 for edge in ['leftBorder','rightBorder','topBorder','bottomBorder']:
  b=x.find('hh:'+edge,ns)
  if b is not None and b.get('type')!='NONE':b.set('color','#D9D9D9')
bfs=h.find('.//hh:borderFills',ns)
base=next(x for x in bfs if x.get('id')==tbl.find('.//hp:tc',ns).get('borderFillIDRef'))
bf=deepcopy(base); new_id=str(max(int(x.get('id')) for x in bfs)+1);bf.set('id',new_id)
fill=bf.find('hc:fillBrush',ns)
if fill is not None:bf.remove(fill)
fill=E.SubElement(bf,'{'+ns['hc']+'}fillBrush');E.SubElement(fill,'{'+ns['hc']+'}winBrush',faceColor='#EAF0F4',hatchColor='#999999',alpha='0')
bfs.append(bf);bfs.set('itemCnt',str(len(bfs)))
for cell in tbl.find('hp:tr',ns):cell.set('borderFillIDRef',new_id);cell.set('header','1')
for x in s.findall('.//hp:linesegarray',ns):x.getparent().remove(x)
for name,tree in [('Contents/header.xml',h),('Contents/section0.xml',s)]:data[name]=E.tostring(tree,encoding='utf-8',xml_declaration=True)
with ZipFile(p,'w',ZIP_DEFLATED) as z:
 for name,blob in data.items():z.writestr(name,blob)
texts=[]
for para in s:
 t=''.join(para.xpath('.//hp:t/text()',namespaces=ns))
 if t:texts.append(t)
Path('D:/compbasic/한글기초/assets/토끼고기요리/본문_입력연습.txt').write_text('\n\n'.join(texts),encoding='utf-8-sig')
