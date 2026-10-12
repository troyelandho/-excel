from pathlib import Path
import pypdfium2 as pdfium
from PIL import Image,ImageOps,ImageDraw
h=Path(__file__).resolve().parent/'视觉检查'
images=[]
for p in sorted(h.glob('sheet_*.pdf')):
 d=pdfium.PdfDocument(str(p));im=d[0].render(scale=1.5).to_pil();im.save(p.with_suffix('.png'))
 small=im.copy();small.thumbnail((780,1100));tile=Image.new('RGB',(800,1140),'white');tile.paste(small,(10,30));ImageDraw.Draw(tile).text((10,10),p.stem,fill='black');images.append(tile)
for start in range(0,len(images),4):
 sel=images[start:start+4];canvas=Image.new('RGB',(1600,2280),'#eeeeee')
 for j,t in enumerate(sel):canvas.paste(t,((j%2)*800,(j//2)*1140))
 canvas.save(h/f'检查拼图_{start//4+1}.png')
print([p.name for p in h.glob('检查拼图_*.png')])
